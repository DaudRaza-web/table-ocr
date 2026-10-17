"""Table OCR tool - local server.

Renders PDF/image pages, crops the user-selected region and sends it to a
vision model running in the local Ollama (WSL), then exports the edited table
to .xlsx. Everything runs on this machine; nothing leaves localhost.
"""
import base64
import io
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --lan listens on all network interfaces so other devices on the local network can connect.
HOST = os.environ.get("TABLE_OCR_HOST", "0.0.0.0" if "--lan" in sys.argv else "127.0.0.1")
PORT = int(os.environ.get("TABLE_OCR_PORT", "8765"))
OLLAMA = os.environ.get("OLLAMA_URL", "http://localhost:11434")
DEFAULT_MODEL = os.environ.get("TABLE_OCR_MODEL", "qwen2.5vl:7b")
TRANSLATE_MODEL = os.environ.get("TABLE_OCR_TRANSLATE_MODEL", "translategemma:12b")
WSL_DISTRO = os.environ.get("TABLE_OCR_WSL_DISTRO", "")  # "" = default distro
WSL = ["wsl.exe"] + (["-d", WSL_DISTRO] if WSL_DISTRO else [])
STATIC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

DOCS = {}  # doc id -> pymupdf.Document
DOCS_LOCK = threading.Lock()

PROMPT = """You are a precise OCR engine for tables.
Transcribe the table in this image into JSON: {"rows": [[cell, cell, ...], ...]}.
Rules:
- One inner array per visual row, top to bottom, including the header row(s).
- Every row must have the same number of cells; use "" for empty cells.
- If a cell spans several columns, put its text in the first cell and "" in the others.
- Copy text exactly as printed (numbers, signs, decimals, currency symbols, %). Do not translate, sum, or invent values.
- Multi-line text inside one cell is joined with a single space.
Return only the JSON."""

SCHEMA = {
    "type": "object",
    "properties": {"rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}}},
    "required": ["rows"],
}


# ---------------------------------------------------------------- Ollama ---

def ollama_request(path, payload=None, timeout=10):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(OLLAMA + path, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())


def ollama_up():
    try:
        ollama_request("/api/version", timeout=2)
        return True
    except Exception:
        return False


def ensure_ollama():
    """Make sure the WSL Ollama is reachable; boot WSL / start `ollama serve` if needed."""
    wsl = WSL
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if sys.platform == "win32":
        # Keep the WSL VM alive while this tool runs (WSL idles out otherwise).
        try:
            subprocess.Popen(wsl + ["-e", "sleep", "infinity"], creationflags=flags,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except OSError:
            pass
    for _ in range(20):
        if ollama_up():
            return True
        time.sleep(0.5)
    if sys.platform == "win32":
        print("Ollama not reachable - starting `ollama serve` inside WSL ...")
        subprocess.Popen(wsl + ["-e", "bash", "-lc", "ollama serve"], creationflags=flags,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for _ in range(40):
            if ollama_up():
                return True
            time.sleep(0.5)
    return False


def parse_rows(text):
    """Model output -> list of rows (list of str). Accepts JSON or a markdown table."""
    text = text.strip()
    m = re.search(r"\{.*\}|\[.*\]", text, re.S)
    if m:
        try:
            obj = json.loads(m.group(0))
            rows = obj.get("rows") if isinstance(obj, dict) else obj
            if isinstance(rows, list):
                return [[("" if c is None else str(c)).strip() for c in (r if isinstance(r, list) else [r])]
                        for r in rows]
        except json.JSONDecodeError:
            pass
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
            continue
        rows.append(cells)
    return rows


def normalize(rows):
    rows = [r for r in rows if any(c.strip() for c in r)]
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]
    # drop columns that are empty everywhere
    keep = [i for i in range(width) if any(r[i].strip() for r in rows)]
    return [[r[i] for i in keep] for r in rows]


# ------------------------------------------------------------- documents ---

def open_document(data, filename):
    ext = os.path.splitext(filename)[1].lower().lstrip(".") or None
    doc = pymupdf.open(stream=data, filetype=ext)
    if not doc.is_pdf:  # images (png/jpg/tiff/...) -> convert so every page renders the same way
        pdf = pymupdf.open("pdf", doc.convert_to_pdf())
        doc.close()
        doc = pdf
    # Bake any /Rotate into the content so every page starts at 0; our own rotation is applied on render.
    for page in doc:
        if page.rotation:
            page.remove_rotation()
    return doc


def norm_rot(rot):
    return int(rot or 0) % 360 // 90 * 90


def unrotate_frac(rect_frac, rot):
    """Selection in the displayed (rotated clockwise by `rot`) page -> fractions of the unrotated page."""
    def pt(u, v):
        return {0: (u, v), 90: (v, 1 - u), 180: (1 - u, 1 - v), 270: (1 - v, u)}[rot]
    u0, v0, u1, v1 = rect_frac
    (ax, ay), (bx, by) = pt(u0, v0), pt(u1, v1)
    return min(ax, bx), min(ay, by), max(ax, bx), max(ay, by)


def render_page(doc, page_no, dpi, rot=0):
    zoom = dpi / 72
    return doc[page_no].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom).prerotate(norm_rot(rot)), alpha=False)


def render_crop(doc, page_no, rect_frac, rot=0):
    rot = norm_rot(rot)
    page = doc[page_no]
    r = page.rect
    x0, y0, x1, y1 = unrotate_frac(rect_frac, rot)
    clip = pymupdf.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height,
                        r.x0 + x1 * r.width, r.y0 + y1 * r.height)
    # Aim for the longest side at ~1600px: sharp enough for small print, small enough for the model.
    longest = max(clip.width, clip.height)
    zoom = max(1.0, min(1600 / longest, 6.0))
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom).prerotate(rot), clip=clip, alpha=False)
    return pix.tobytes("png")


# ------------------------------------------------------- orientation check --

def orientation_from_text(page):
    """Digital PDFs: the direction of the text lines says exactly how far the page is turned."""
    votes = {0: 0, 90: 0, 180: 0, 270: 0}
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            dx, dy = line["dir"]
            n = sum(len(s["text"].strip()) for s in line["spans"])
            if abs(dx) >= abs(dy):
                votes[0 if dx > 0 else 180] += n
            else:  # text running down the page needs a 270° (anticlockwise) turn, and vice versa
                votes[270 if dy > 0 else 90] += n
    total = sum(votes.values())
    if total < 40:
        return None
    best = max(votes, key=votes.get)
    return best if votes[best] >= 0.6 * total else None


def is_sideways(page):
    """Scans: lines of text make dark/light stripes. Stronger stripes across columns than rows means the
    text runs vertically, i.e. the page is turned 90° one way or the other. None = can't tell."""
    zoom = 600 / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csGRAY, alpha=False)
    w, h, s = pix.width, pix.height, pix.samples
    rows, cols = [0] * h, [0] * w
    for y in range(h):
        line = s[y * pix.stride: y * pix.stride + w]
        for x, v in enumerate(line):
            if v < 128:
                rows[y] += 1
                cols[x] += 1
    if sum(rows) < 0.002 * w * h:
        return None  # nearly blank

    def stripes(profile):
        return sum((a - b) ** 2 for a, b in zip(profile, profile[1:])) / max(1, len(profile))

    r, c = stripes(rows) / max(1, w), stripes(cols) / max(1, h)  # normalise by line length
    if c > 1.5 * r:
        return True
    if r > 1.5 * c:
        return False
    return None


_TESSERACT = None


def tesseract_cmd():
    """Tesseract on PATH, else (on Windows) inside WSL. [] if unavailable. Vision LLMs can't reliably tell
    upright from upside-down; Tesseract's orientation detector can, for Latin and Arabic scripts alike."""
    global _TESSERACT
    if _TESSERACT is None:
        _TESSERACT = []
        if shutil.which("tesseract"):
            _TESSERACT = ["tesseract"]
        elif sys.platform == "win32":
            try:
                r = subprocess.run(WSL + ["-e", "which", "tesseract"], capture_output=True, timeout=20,
                                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                if r.returncode == 0:
                    _TESSERACT = WSL + ["-e", "tesseract"]
            except (OSError, subprocess.TimeoutExpired):
                pass
    return _TESSERACT


def orientation_from_tesseract(png):
    cmd = tesseract_cmd()
    if not cmd:
        return None
    try:
        r = subprocess.run(cmd + ["stdin", "stdout", "--psm", "0"], input=png, capture_output=True, timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired):
        return None
    out = r.stdout.decode(errors="replace")
    rot = re.search(r"Rotate:\s*(\d+)", out)
    conf = re.search(r"Orientation confidence:\s*([\d.]+)", out)
    if rot and conf and float(conf.group(1)) >= 1.0:
        return norm_rot(rot.group(1))
    return None


def detect_orientation(doc, page_no):
    """-> (clockwise turn that makes the page upright, how it was decided)."""
    with DOCS_LOCK:  # MuPDF isn't thread-safe; Tesseract runs outside the lock
        page = doc[page_no]
        rot = orientation_from_text(page)
        if rot is not None:
            return rot, "text"
        zoom = 2000 / max(page.rect.width, page.rect.height)
        png = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csGRAY, alpha=False).tobytes("png")
    rot = orientation_from_tesseract(png)
    if rot is not None:
        return rot, "tesseract"
    with DOCS_LOCK:
        sideways = is_sideways(doc[page_no])
    if sideways:  # no Tesseract (or unsure): we know it's on its side, not which way
        return 90, "guess"
    return 0, "unknown"


# ------------------------------------------------------------- translation --

# code -> (name used in the prompt, label shown in the UI)
LANGUAGES = {
    "ur": ("Urdu", "Urdu"),
    "pa-Arab": ("Punjabi", "Punjabi (Shahmukhi)"),
    "pa-Guru": ("Punjabi", "Punjabi (Gurmukhi)"),
    "sd-Arab": ("Sindhi", "Sindhi"),
    "en": ("English", "English"),
}
# Letters used in Sindhi but not in Urdu / Shahmukhi Punjabi.
SINDHI_LETTERS = set("ٻٽٿڀڃڄڇڊڌڍڏڙڦڪڳڱ")
PUNJABI_WORDS = {"دا", "دی", "دے", "نوں", "وچ", "توں", "تے", "ایہہ", "اوہ", "سی", "ہیگا", "کیہ", "جی", "اسیں", "تسیں"}
URDU_WORDS = {"کا", "کی", "کے", "کو", "میں", "سے", "ہے", "ہیں", "تھا", "یہ", "وہ", "اور", "نے"}


def detect_language(text):
    if re.search(r"[਀-੿]", text):
        return "pa-Guru"
    if any(ch in SINDHI_LETTERS for ch in text):
        return "sd-Arab"
    if not re.search(r"[؀-ۿ]", text):
        return "en"
    # Shahmukhi Punjabi shares Urdu's letters; tell them apart by common grammar words.
    words = re.findall(r"[؀-ۿݐ-ݿ]+", text)
    punjabi = sum(w in PUNJABI_WORDS for w in words) + 3 * text.count("ݨ")
    urdu = sum(w in URDU_WORDS for w in words)
    return "pa-Arab" if punjabi >= 2 and punjabi > urdu else "ur"


def translate_prompt(src, tgt, text):
    s_name, t_name = LANGUAGES[src][0], LANGUAGES[tgt][0]
    # Prompt format TranslateGemma was trained on (two blank lines before the text).
    return (f"You are a professional {s_name} ({src}) to {t_name} ({tgt}) translator. "
            f"Your goal is to accurately convey the meaning and nuances of the original {s_name} text "
            f"while adhering to {t_name} grammar, vocabulary, and cultural sensitivities.\n"
            f"Produce only the {t_name} translation, without any additional explanations or commentary. "
            f"Please translate the following {s_name} text into {t_name}:\n\n\n{text}")


# Reading and translating an image in one step makes the model invent content, so selections are
# transcribed first (it reads Urdu/Sindhi script well) and the text is then translated normally.
def transcribe_prompt(src):
    s_name = LANGUAGES[src][1] if src in LANGUAGES else "Urdu, Punjabi (Shahmukhi or Gurmukhi) or Sindhi"
    return (f"Transcribe all text in this image exactly as written, in its original script and language "
            f"(it may be {s_name}). Do not translate. Keep line and paragraph breaks. "
            f"Output only the transcribed text.")


def clean_transcription(text):
    text = re.sub(r"</?(start|end)_of_image>", "", text)
    lines = text.strip().splitlines()
    # The model sometimes prefixes a bare language label ("Punjabi") on its own line.
    if lines and lines[0].strip().lower() in {"urdu", "punjabi", "sindhi", "shahmukhi", "gurmukhi"}:
        lines = lines[1:]
    return "\n".join(lines).strip()


def chunk_text(text, limit=1500):
    """Split on paragraph (then sentence) boundaries so each request stays well inside the model's comfort zone."""
    chunks, cur = [], ""
    for para in re.split(r"\n\s*\n", text.strip()):
        pieces = [para] if len(para) <= limit else re.split(r"(?<=[.!?۔؟])\s+", para)
        for piece in pieces:
            if cur and len(cur) + len(piece) + 2 > limit:
                chunks.append(cur)
                cur = ""
            cur = f"{cur}\n\n{piece}" if cur else piece
    if cur:
        chunks.append(cur)
    return chunks


def ollama_generate(model, prompt, image=None):
    msg = {"role": "user", "content": prompt}
    if image:
        msg["images"] = [base64.b64encode(image).decode()]
    res = ollama_request("/api/chat", {
        "model": model, "stream": False, "keep_alive": "15m",
        "options": {"temperature": 0, "num_ctx": 8192, "num_predict": 4096},
        "messages": [msg],
    }, timeout=900)
    return res.get("message", {}).get("content", "").strip()


# ------------------------------------------------------------ excel export --

NUM_RE = re.compile(r"^\(?-?[$€£]?\s?\d{1,3}(,\d{3})*(\.\d+)?\)?%?$|^\(?-?[$€£]?\s?\d+(\.\d+)?\)?%?$")


def to_cell_value(s):
    t = s.strip()
    if not t or not NUM_RE.match(t) or (len(t) > 1 and t[0] == "0" and t[1].isdigit()):
        return s, None
    neg = t.startswith("(") and t.endswith(")") or t.lstrip("(").startswith("-")
    pct = t.endswith("%")
    core = re.sub(r"[()$€£%,\s-]", "", t)
    try:
        v = float(core)
    except ValueError:
        return s, None
    if neg:
        v = -v
    decimals = len(core.split(".")[1]) if "." in core else 0
    frac = ("." + "0" * decimals) if decimals else ""
    if pct:
        return v / 100, "0" + frac + "%"
    sym = next((ch for ch in "$€£" if ch in t), "")
    base = "#,##0" if "," in t or sym else "0"
    if not decimals:
        v = int(v)
    fmt = sym + base + frac  # keep printed separators/decimals (e.g. 1.50, $1,200)
    return v, (None if fmt == "0" else fmt)


def build_xlsx(title, rows, header):
    wb = Workbook()
    ws = wb.active
    ws.title = (re.sub(r"[\\/*?:\[\]]", "", title)[:31] or "Table")
    width = max((len(r) for r in rows), default=1)
    thin = Side(style="thin", color="A0A0A0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    start = 1
    if title.strip():
        ws.cell(row=1, column=1, value=title.strip()).font = Font(bold=True, size=14)
        if width > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
        start = 3
    for i, row in enumerate(rows):
        for j in range(width):
            text = row[j] if j < len(row) else ""
            is_head = header and i == 0
            value, fmt = (text, None) if is_head else to_cell_value(text)
            c = ws.cell(row=start + i, column=j + 1, value=value)
            c.border = border
            if fmt:
                c.number_format = fmt
            if is_head:
                c.font = Font(bold=True)
                c.fill = PatternFill("solid", fgColor="D9E1F2")
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j in range(width):
        longest = max((len(r[j]) if j < len(r) else 0 for r in rows), default=8)
        ws.column_dimensions[get_column_letter(j + 1)].width = min(max(10, longest + 2), 60)
    if header and rows:
        ws.freeze_panes = ws.cell(row=start + 1, column=1)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------ HTTP ---

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("%s\n" % (fmt % args))

    def send(self, code, body, ctype="application/json", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body).encode()
        elif isinstance(body, str):
            body = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def body(self):
        return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def get_doc(self, doc_id):
        with DOCS_LOCK:
            doc = DOCS.get(doc_id)
        if doc is None:
            raise KeyError("Document not found - please reopen the file.")
        return doc

    def do_GET(self):
        url = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(url.query).items()}
        try:
            if url.path in ("/", "/index.html"):
                with open(os.path.join(STATIC, "index.html"), "rb") as f:
                    return self.send(200, f.read(), "text/html; charset=utf-8")
            if url.path == "/api/models":
                try:
                    tags = ollama_request("/api/tags", timeout=5)
                except Exception as e:
                    return self.send(200, {"ok": False, "error": f"Ollama not reachable at {OLLAMA}: {e}",
                                           "models": [], "default": DEFAULT_MODEL,
                                           "translate_default": TRANSLATE_MODEL})
                names = [m["name"] for m in tags.get("models", [])]
                return self.send(200, {"ok": True, "models": names, "default": DEFAULT_MODEL,
                                       "translate_default": TRANSLATE_MODEL})
            if url.path == "/api/page":
                doc = self.get_doc(q["id"])
                with DOCS_LOCK:
                    png = render_page(doc, int(q["n"]), float(q.get("dpi", 150)), q.get("rot")).tobytes("png")
                return self.send(200, png, "image/png")
            if url.path == "/api/orient":
                doc = self.get_doc(q["id"])
                rot, how = detect_orientation(doc, int(q["n"]))
                return self.send(200, {"rotation": rot, "method": how})
            self.send(404, {"error": "not found"})
        except Exception as e:
            self.send(400, {"error": str(e)})

    def do_POST(self):
        url = urlparse(self.path)
        try:
            if url.path == "/api/upload":
                name = unquote(self.headers.get("X-Filename", "file.pdf"))
                doc = open_document(self.body(), name)
                doc_id = uuid.uuid4().hex
                with DOCS_LOCK:
                    DOCS[doc_id] = doc
                return self.send(200, {"id": doc_id, "pages": doc.page_count, "name": name})

            if url.path == "/api/extract":
                req = json.loads(self.body())
                doc = self.get_doc(req["id"])
                with DOCS_LOCK:
                    png = render_crop(doc, int(req["page"]), req["rect"], req.get("rot"))
                model = req.get("model") or DEFAULT_MODEL
                t0 = time.time()
                res = ollama_request("/api/chat", {
                    "model": model,
                    "stream": False,
                    "format": SCHEMA,
                    "keep_alive": "15m",
                    "options": {"temperature": 0, "num_ctx": 16384, "num_predict": 8192},
                    "messages": [{"role": "user", "content": PROMPT,
                                  "images": [base64.b64encode(png).decode()]}],
                }, timeout=900)
                text = res.get("message", {}).get("content", "")
                rows = normalize(parse_rows(text))
                if not rows:
                    return self.send(200, {"ok": False, "error": "The model did not return a table. Try a tighter selection.",
                                           "raw": text})
                return self.send(200, {"ok": True, "rows": rows, "seconds": round(time.time() - t0, 1),
                                       "model": model})

            if url.path == "/api/translate":
                req = json.loads(self.body())
                model = req.get("model") or TRANSLATE_MODEL
                tgt = req.get("target") or "en"
                src = req.get("source") or "auto"
                t0 = time.time()
                if req.get("rect"):  # selected region of the page: transcribe first, then translate
                    doc = self.get_doc(req["id"])
                    with DOCS_LOCK:
                        png = render_crop(doc, int(req["page"]), req["rect"], req.get("rot"))
                    text = clean_transcription(ollama_generate(model, transcribe_prompt(src), png))
                    if not text:
                        return self.send(200, {"ok": False, "error": "No text found in the selection."})
                else:
                    text = (req.get("text") or "").strip()
                    if not text:
                        return self.send(200, {"ok": False, "error": "Nothing to translate."})
                if src == "auto":
                    src = detect_language(text)
                if src == tgt:
                    return self.send(200, {"ok": False, "original": text,
                                           "error": f"Text looks like {LANGUAGES[src][1]} already - pick the source language."})
                out = "\n\n".join(ollama_generate(model, translate_prompt(src, tgt, c)) for c in chunk_text(text))
                return self.send(200, {"ok": True, "translation": out, "original": text, "source": src,
                                       "seconds": round(time.time() - t0, 1), "model": model})

            if url.path == "/api/export":
                req = json.loads(self.body())
                title = req.get("title", "")
                data = build_xlsx(title, req.get("rows", []), bool(req.get("header", True)))
                fname = (re.sub(r'[\\/:*?"<>|]+', "_", title).strip() or "table") + ".xlsx"
                return self.send(200, data,
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 {"Content-Disposition": f'attachment; filename="{fname}"'})

            if url.path == "/api/close":
                req = json.loads(self.body())
                with DOCS_LOCK:
                    doc = DOCS.pop(req.get("id"), None)
                if doc:
                    doc.close()
                return self.send(200, {"ok": True})
            self.send(404, {"error": "not found"})
        except urllib.error.HTTPError as e:
            detail = e.read().decode(errors="replace")
            self.send(502, {"error": f"Ollama error {e.code}: {detail}"})
        except Exception as e:
            self.send(400, {"error": f"{type(e).__name__}: {e}"})


def lan_ips():
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    return sorted(ip for ip in ips if not ip.startswith(("127.", "169.254.", "172.")))  # skip loopback/WSL/Docker


def main():
    print(f"Checking Ollama at {OLLAMA} ...")
    print("Ollama OK" if ensure_ollama() else "WARNING: Ollama not reachable; extraction will fail until it is up.")
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://127.0.0.1:{PORT}/"
    print(f"Table OCR running at {url}  (Ctrl+C to stop)")
    if HOST != "127.0.0.1":
        print("Open from other devices on your network at:")
        for ip in lan_ips():
            print(f"    http://{ip}:{PORT}/")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
