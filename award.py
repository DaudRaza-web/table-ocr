"""Award document extraction.

Award documents (Land Acquisition Act awards) are typed English. Their "detail of land" table has
District | Tehsil | Locality/Mouza (or Chack) | Sqr. No | Khasra (Killa) No. | K | M and can run over
several pages; the total is printed at the end ("264K - 10M" or a "Total:- 542 | 02" row).

Pipeline: find the table's ruled grid in code, crop its right-hand columns on each page and read them
with the table extractor (the same prompt as the Table tab), then build the land list in code:
the square number is carried down over blank rows and joined to the khasra (square 2, khasra 19/1 ->
"2/19/1"), plots printed in the Sqr column without a square ("116/1") are kept as they are, and the
stated total closes the table. Individual K/M areas are not reported; only the document's total is.
"""
import csv
import io
import json
import os
import re
import subprocess

import pymupdf

DPI = 300


# ------------------------------------------------------------------ helpers --

def _gray(page, rot=0, dpi=DPI):
    z = dpi / 72
    return page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), colorspace=pymupdf.csGRAY, alpha=False)


def _dark(pix, thr=150):
    w, h, s, st = pix.width, pix.height, pix.samples, pix.stride
    return w, h, [bytes(1 if v < thr else 0 for v in s[y * st: y * st + w]) for y in range(h)]


def _group(idxs, gap=4):
    out = []
    for i in idxs:
        if out and i - out[-1][-1] <= gap:
            out[-1].append(i)
        else:
            out.append([i])
    return [(g[0], g[-1]) for g in out]


def tesseract(cmd, png, psm=6):
    """-> list of words {text, left, top, width, height, conf}."""
    r = subprocess.run(cmd + ["stdin", "stdout", "-l", "eng", "--psm", str(psm), "tsv"], input=png,
                       capture_output=True, timeout=120, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    rows = csv.DictReader(io.StringIO(r.stdout.decode(errors="replace")), delimiter="\t", quoting=csv.QUOTE_NONE)
    out = []
    for w in rows:
        if w.get("level") == "5" and (w.get("text") or "").strip():
            out.append({"text": w["text"].strip(), "left": int(w["left"]), "top": int(w["top"]),
                        "width": int(w["width"]), "height": int(w["height"]), "conf": float(w["conf"])})
    return out


# --------------------------------------------------------------------- grid --

def find_table(pix):
    """Ruled grid of the land table: vertical rules that run a long way down the page.
    -> {"v": [x...], "cols": last five rules, "y0", "y1", "size"} in pixels, or None."""
    w, h, D = _dark(pix, thr=215)  # table rules in some scans are light grey
    # longest vertical dark run per column (1 px of horizontal wobble allowed)
    best = [0] * w
    edge = int(0.015 * w)  # the scanned paper's own edge is not a rule
    for x in range(edge, w - edge):
        run = top = 0
        for y in range(h):
            if D[y][x] or D[y][x - 1] or D[y][x + 1]:
                run += 1
                if run > top:
                    top = run
            else:
                run = 0
        best[x] = top
    cand = [x for x in range(w) if best[x] > 0.12 * h]
    vs = [(a + b) // 2 for a, b in _group(cand, gap=6)]
    if len(vs) < 5:
        return None
    # The last four columns are always Sqr | Khasra | K | M (the District/Tehsil/Mouza columns on the
    # left are not ruled all the way down in every layout).
    right = vs[-5:]
    starts, ends = [], []
    for x in right:
        # the rule's longest stretch, bridging small breaks in the scan (up to 2.5% of the page)
        segs, s0, last = [], None, None
        for y in range(h):
            if D[y][x] or D[y][x - 1] or D[y][x + 1]:
                if s0 is None or y - last > 0.025 * h:
                    if s0 is not None:
                        segs.append((s0, last))
                    s0 = y
                last = y
        if s0 is not None:
            segs.append((s0, last))
        a, b = max(segs, key=lambda s: s[1] - s[0]) if segs else (0, 0)
        starts.append(a)
        ends.append(b)
    starts.sort()
    ends.sort()
    return {"v": vs, "cols": right, "y0": starts[2], "y1": ends[2], "size": (w, h)}


# -------------------------------------------------------------- read a page --

def read_table(page, ask, rot=0):
    """Land table cells on one page via the table extractor.
    `ask(png) -> list of rows (list of str)`. -> rows, or None when the page has no ruled table."""
    g = find_table(_gray(page, rot, dpi=200))
    if not g:
        return None
    w, h = g["size"]
    v, c = g["v"], g["cols"]
    # one column more on the left (Locality/Mouza) so a "Total" label printed there is read too
    left = v[-6] if len(v) >= 6 else c[0]
    r = page.rect
    rect = pymupdf.Rect((left / w - 0.005) * r.width, max(0, g["y0"] / h - 0.005) * r.height,
                        (c[4] / w + 0.005) * r.width, min(1, g["y1"] / h + 0.035) * r.height)
    z = 1800 / max(rect.width, rect.height)
    png = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), clip=rect, alpha=False).tobytes("png")
    return ask(png)


# ------------------------------------------------------------- build table --

NUM = re.compile(r"^\d{1,4}$")
KHASRA = re.compile(r"^\d{1,4}(/\d{1,4})*$")
TOTAL_KM = re.compile(r"(\d{1,5})\s*K\.?\s*[-–—~,]?\s*(\d{1,2})\s*M", re.I)
TOTAL_ROW = re.compile(r"total\D*(\d{1,5})\D+(\d{1,2})\b", re.I)


def _clean(t):
    t = (t or "").strip().strip(".,;:|[](){}_'\"`~-—").replace(" ", "")
    return t.replace("O", "0").replace("o", "0").replace("l", "1").replace("I", "1").replace("\\", "/")


def build(pages):
    """[(page_no, rows)] from the table extractor -> {"rows": [{"khasra", "page"}], "total", "issues"}.

    Per row: K and M are the last two plain numbers; before them come [square] and khasra. The table
    extractor sometimes leaves out blank cells, so a lone value counts as the khasra unless it sits in
    the Sqr column (a blank cell after it) or is above 30 with no square (killas run 1-25, sometimes
    26-27) - then it is a plot numbered without a square and kept as printed."""
    rows, issues, total, sqr = [], [], None, ""
    seen = {}
    for page_no, cells in pages:
        for raw in cells:
            joined = " ".join(x for x in raw if x)
            t = TOTAL_KM.search(joined) or TOTAL_ROW.search(joined)
            if t:
                total = (int(t.group(1)), int(t.group(2)))
                return {"rows": rows, "total": total, "issues": issues}
            # headings and place names are text; only numbers matter here
            # (a number misread with l/I/o/O in it is still a number)
            cl = ["" if re.search(r"[A-Za-z]", x) and not re.fullmatch(r"[\d/lIoO\s.]+", x) else _clean(x) for x in raw]
            idx = [i for i, x in enumerate(cl) if x]
            if not idx:
                continue
            nums = [i for i in idx if NUM.match(cl[i])]
            if len(nums) >= 2 and nums[-1] == idx[-1] and nums[-2] == idx[-2] and len(idx) >= 3:
                pre = cl[:nums[-2]]
            else:
                pre = cl  # no K/M read on this line
                issues.append(f"Page {page_no}: no K/M area on line '{joined}' - check this row")
            vals = [(i, x) for i, x in enumerate(pre) if x]
            if not vals:
                continue
            plot = False
            if len(vals) >= 2:
                sq, kh = vals[0][1], vals[1][1]
                if len(vals) > 2:
                    issues.append(f"Page {page_no}: extra values on line '{joined}' - check this row")
            else:
                i, kh = vals[0]
                sq = ""
                in_sqr_col = len(pre) >= 2 and i == len(pre) - 2 and pre[-1] == ""
                first = re.match(r"\d+", kh)
                plot = in_sqr_col or (first is not None and int(first.group(0)) > 30)
            kh = kh.strip("/")
            if not KHASRA.match(kh) or (sq and not NUM.match(sq)):
                issues.append(f"Page {page_no}: unclear line '{joined}' - check the khasra number")
            if sq:
                sqr = sq.lstrip("0") or sq
            name = kh if plot or not sqr else f"{sqr}/{kh}"
            if not plot and not sqr and KHASRA.match(kh):
                issues.append(f"Page {page_no}: khasra {kh} has no square number above it")
            if name in seen:
                issues.append(f"Page {page_no}: khasra {name} appears twice (also on page {seen[name]}) - "
                              "a line may have been read twice")
            seen.setdefault(name, page_no)
            rows.append({"khasra": name, "page": page_no})
    return {"rows": rows, "total": total, "issues": issues}


# ------------------------------------------------------------ award header --

PLACE_SCHEMA = {"type": "object", "properties": {"district": {"type": "string"}, "tehsil": {"type": "string"},
                "locality": {"type": "string"}}, "required": ["district", "tehsil", "locality"]}
PLACE_PROMPT = ("This is the top-left of a typed land table. Its headings are District, Tehsil and Locality/Mouza "
                "(or Chack/Chak/Village). Copy the value printed under each heading exactly, as JSON "
                '{"district": "", "tehsil": "", "locality": ""}. Use "" for a heading that is not there.')


def read_place(page, ask_json, rot=0):
    """District / Tehsil / Locality from the table's top-left. `ask_json(prompt, png, schema) -> str`."""
    g = find_table(_gray(page, rot, dpi=200))
    if not g:
        return {}
    w, h = g["size"]
    r = page.rect
    y = g["y0"] / h
    rect = pymupdf.Rect(0.03 * r.width, max(0, y - 0.05) * r.height, (g["cols"][0] / w) * r.width, min(1, y + 0.07) * r.height)
    z = 1400 / max(rect.width, 1)
    png = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), clip=rect, alpha=False).tobytes("png")
    try:
        d = json.loads(ask_json(PLACE_PROMPT, png, PLACE_SCHEMA))
    except (ValueError, TypeError):
        return {}
    return {"district": d.get("district", "").strip(), "tehsil": d.get("tehsil", "").strip(),
            "mouza": d.get("locality", "").strip()}


AWARD_NO = re.compile(r"Award\s*No\.?\s*:?\s*(\d{1,3})\s*(?:[/-]\s*(\d{4}))?", re.I)


def award_number(path, page=None, cmd=None):
    """"12/2015" from the file name ("Award No. 12-2015 ..."), else from the page header by Tesseract."""
    m = AWARD_NO.search(os.path.basename(path))
    if m and m.group(2):
        return f"{int(m.group(1))}/{m.group(2)}"
    if page is not None and cmd:
        z = DPI / 72
        r = page.rect
        png = page.get_pixmap(matrix=pymupdf.Matrix(z, z), clip=pymupdf.Rect(0, 0, r.width, r.height * 0.15),
                              colorspace=pymupdf.csGRAY, alpha=False).tobytes("png")
        t = AWARD_NO.search(" ".join(wd["text"] for wd in tesseract(cmd, png, psm=6)))
        if t and t.group(2):
            return f"{int(t.group(1))}/{t.group(2)}"
    return f"{int(m.group(1))}" if m else ""


# ---------------------------------------------------------------- document --

def extract(path, ask_table, ask_json, cmd=None, max_pages=None, progress=None):
    """Whole award PDF -> {"award", "place", "rows": [{"khasra", "page"}], "total", "issues", "pages"}.
    `ask_table(png) -> rows`; `ask_json(prompt, png, schema) -> str`; `cmd` is Tesseract (award number)."""
    doc = pymupdf.open(path)
    n_pages = min(doc.page_count, max_pages or doc.page_count)
    pages, used, place, issues = [], [], {}, []
    for n in range(n_pages):
        if progress:
            progress(n, n_pages)
        cells = read_table(doc[n], ask_table)
        res = build(pages + [(n + 1, cells or [])])
        if not res["rows"] and res["total"] is None:
            if used:
                issues.append(f"Page {n + 1}: no land table here although the total was not reached yet")
                break
            continue  # the land table has not started
        pages.append((n + 1, cells or []))
        used.append(n + 1)
        if not place:
            place = read_place(doc[n], ask_json)
        if res["total"] is not None:
            break  # the total closes the land table
    res = build(pages)
    issues += res["issues"]
    if not used:
        issues.append("No land table (Sqr. No / Khasra No. / K / M) was found in the document")
    elif res["total"] is None:
        issues.append("No total K-M found at the end of the land table" +
                      (f" (only the first {n_pages} pages were read)" if max_pages else ""))
    for k, label in (("district", "District"), ("tehsil", "Tehsil"), ("mouza", "Locality/Mouza")):
        if not place.get(k):
            issues.append(f"{label} could not be read from the table - fill it in")
    award = award_number(path, doc[0] if doc.page_count else None, cmd)
    if not award:
        issues.append("Award number not found in the file name or the page header")
    return {"award": award, "place": place, "rows": res["rows"], "total": res["total"], "issues": issues,
            "pages": used}


# ------------------------------------------------------------------- excel --

HEADERS = ["S.No.", "Award No", "District", "Tehsil", "Locality/Mouza", "Khasra No.", "K", "M", "Page"]


def table_rows(award, place, rows, total):
    """Rows of the land table under HEADERS; the K and M columns stay blank except on the closing Total row."""
    out = [[i, award, place.get("district", ""), place.get("tehsil", ""), place.get("mouza", ""), r["khasra"], "", "",
            r["page"]] for i, r in enumerate(rows, 1)]
    k, m = total or ("", "")
    return out + [["", "", "", "", "", "Total", k, m, ""]]


def build_xlsx(rows, info, issues):
    """rows: table_rows() (possibly edited); info: [(item, detail)] for the Review sheet; issues: [text].
    -> .xlsx bytes"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    def num(v):
        s = str(v if v is not None else "").strip()
        return int(s) if re.fullmatch(r"\d{1,6}", s) else s

    wb = Workbook()
    ws = wb.active
    ws.title = "Award Land"
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    head = PatternFill("solid", fgColor="DDE6F0")
    ws.append(HEADERS)
    for c in ws[1]:
        c.font, c.fill, c.border = Font(bold=True), head, box
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        r = list(r) + [""] * (len(HEADERS) - len(r))
        # khasra numbers stay text ("2/19/1"); S.No., K, M and Page are numbers
        ws.append([num(r[0]), r[1], r[2], r[3], r[4], str(r[5]), num(r[6]), num(r[7]), num(r[8])])
        bold = str(r[5]).strip().lower() == "total"
        for c in ws[ws.max_row]:
            c.border = box
            c.alignment = Alignment(horizontal="center")
            if bold:
                c.font = Font(bold=True)
    for col, wdt in zip("ABCDEFGHI", (7, 11, 16, 16, 20, 14, 7, 7, 7)):
        ws.column_dimensions[col].width = wdt
    ws.freeze_panes = "A2"

    rv = wb.create_sheet("Review")
    rv.append(["Item", "Detail"])
    for c in rv[1]:
        c.font, c.fill = Font(bold=True), head
    for item in info:
        rv.append(list(item))
    rv.append([])
    rv.append(["Check", "Issue"])
    for c in rv[rv.max_row]:
        c.font = Font(bold=True)
    for it in issues or ["Nothing looked off in the extraction"]:
        rv.append(["!" if issues else "OK", it])
    rv.column_dimensions["A"].width = 26
    rv.column_dimensions["B"].width = 110
    if issues:
        rv.sheet_properties.tabColor = "E0A030"  # draws the eye to the Review tab
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
