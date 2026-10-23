"""Mutation register ("Register Dakhil Kharij") extraction.

Same job as the cloud Mutation Register Extractor, done with local models. Small vision models can't
reliably work out which of 17 Urdu-headed columns is which, so the layout is found in code instead:
the register is a printed, ruled form, and the row of column numbers (١ … ١٧) under its top border
gives every column's position. The cells that matter are then cut out and each one is read by the
local vision model on its own (reading one small cell is something these models do well), and the
values are parsed with plain rules:

  column 1   Award / mutation number
  column 2   Khata / Khatuni numbers ("440 تا 444" = 440 to 444)
  column 14  area, Kanal-Marla-Sarsai (or sq ft)
  column 17  total area with the Katat (plot count) printed under it
  header     ضلع / تحصیل / محال (District / Tehsil / Mouza)
"""
import contextlib
import re

import pymupdf

RENDER_LONG = 2000  # px, longest side used for finding the grid
INK_MIN = 0.004     # a cell with less ink than this is treated as empty

PROMPT_PLAIN = ("Transcribe all text in this image exactly as written, top to bottom, one line per printed line. "
                "Keep digits, hyphens and slashes exactly. Output only the text.")
PROMPT_KHATA = ("This image is one cell from an Urdu land register. It holds a vertical list of khata numbers. "
                "Between two numbers there may be the short Urdu word تا (it means 'to' and marks a range). "
                "Write each printed line on its own line, top to bottom. Write numbers with Western digits. "
                "Wherever the word تا appears, write the word TO on that line instead. Output only the lines.")

EASTERN = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
KMS = re.compile(r"(\d+)\s*[-–]\s*(\d+)\s*[-–]\s*(\d+(?:\.\d+)?)")


def _norm(text):
    return (text or "").translate(EASTERN).replace("<start_of_image>", "").replace("<end_of_image>", "")


def parse_award(text):
    m = re.search(r"\b\d{1,7}\b", _norm(text))
    return m.group(0) if m else ""


def parse_kms(text):
    m = KMS.search(_norm(text))
    return (m.group(1), m.group(2), m.group(3)) if m else None


def parse_total_katat(text):
    """Column 17: the total area (K-M-S) and the small whole number printed under it (the Katat)."""
    t = _norm(text)
    m = KMS.search(t)
    if not m:
        return None, ""
    after = t[m.end():]
    k = re.search(r"(?<![\d/.-])(\d{1,4})(?![\d/.-])", after)
    return m.group(0), (k.group(1) if k else "")


def _plain_numbers(text):
    """Whole numbers in a transcription, skipping fractions (26/1601), CNIC-style IDs and K-M-S areas."""
    t = KMS.sub(" ", _norm(text))
    t = re.sub(r"\d+\s*/\s*\d+", " ", t)
    t = re.sub(r"\d+(?:-\d+){1,}", " ", t)
    return [int(n) for n in re.findall(r"(?<!\d)(\d{1,6})(?!\d)", t)]


def parse_khata(text):
    """'133 / 440 / TO / 444' -> [(133, None), (440, 444)]."""
    items, pending_to = [], False
    t = re.sub(r"\bto\b|تا", " TO ", _norm(text), flags=re.I)
    t = re.sub(r"\d+\s*/\s*\d+", " ", t)
    for tok in re.findall(r"TO|\d{1,6}", t):
        if tok == "TO":
            pending_to = bool(items) and items[-1][1] is None
            continue
        n = int(tok)
        if pending_to:
            items[-1] = (items[-1][0], n)
            pending_to = False
        else:
            items.append((n, None))
    return items


def format_khata(items):
    return ", ".join(str(a) if b is None or b == a else f"{a} to {b}" for a, b in items)


def _dark_rows(pix):
    w, h, s, st = pix.width, pix.height, pix.samples, pix.stride
    return w, h, [bytes(1 if v < 200 else 0 for v in s[y * st: y * st + w]) for y in range(h)]


def _group(idxs, gap=3):
    out = []
    for i in idxs:
        if out and i - out[-1][-1] <= gap:
            out[-1].append(i)
        else:
            out.append([i])
    return [(g[0] + g[-1]) // 2 for g in out]


def find_grid(page, rot=0):
    """Locate the register table on an (upright after `rot`) page.
    Returns {"cols": [(x0, x1) for column 1..17], "rows": [(y0, y1) data bands], "top": y, "size": (w, h)}
    in fractions of the rotated page, or None if no 17-column ruled table is found."""
    z = RENDER_LONG / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), colorspace=pymupdf.csGRAY, alpha=False)
    w, h, D = _dark_rows(pix)

    # Horizontal rules: rows that are mostly dark (OR-ing 3 rows tolerates slight skew and thin lines).
    counts = []
    for y in range(h - 2):
        a, b, c = D[y], D[y + 1], D[y + 2]
        counts.append(sum(1 for i in range(w) if a[i] or b[i] or c[i]))
    lines = _group([y for y, n in enumerate(counts) if n > 0.5 * w])
    if len(lines) < 3:
        return None

    def verticals(y0, y1, frac=0.9):
        y0, y1 = y0 + 3, y1 - 3
        if y1 - y0 < 6:
            return []
        need = frac * (y1 - y0)
        xs = [x for x in range(w - 1) if sum(1 for y in range(y0, y1) if D[y][x] or D[y][x + 1]) >= need]
        return _group(xs)

    bands = [(a, b, verticals(a, b)) for a, b in zip(lines, lines[1:])]
    # The column-number row: the first band crossed by 18 rules (17 columns). It's short, so be lenient.
    num = next((i for i, (a, b, v) in enumerate(bands) if 16 <= len(v) <= 20), None)
    if num is None:
        return None
    v = bands[num][2]
    if len(v) != 18:  # a rule was missed or doubled: keep the 18 that best match the outer borders
        return None
    cols = [(v[17 - k] / w, v[18 - k] / w) for k in range(1, 18)]  # column 1 is rightmost

    # Header: after the number row comes a row of merged group headings (few rules) and then the
    # column titles. Data rows are the ruled bands after that which still carry most column rules.
    i = num + 1
    if i < len(bands) and len(bands[i][2]) < 12:
        i += 1          # group-heading row
    i += 1              # column-title row
    rows = [(a / h, b / h) for a, b, vv in bands[i:] if len(vv) >= 12 and (b - a) > 0.02 * h]
    return {"cols": cols, "rows": rows, "top": bands[num][0] / h, "size": (w, h)}


def read_page(page, rot, ask, reader_model, khata_model, lock=None):
    """Read every mutation entry on one register page.
    `ask(model, prompt, png) -> str` calls the local vision model. `lock` (optional) is held while
    rendering — MuPDF isn't thread-safe — but not during the slow model calls. Returns a dict:
    {"ok": bool, "error": str, "records": [{award, khata, katat, kanal, marla, sarsai, continuation,
     confidence, notes, cells}], "grid": {...}} — cells/grid (fractions) let the UI outline what was read."""
    lock = lock or contextlib.nullcontext()
    with lock:
        grid = find_grid(page, rot)
    if not grid:
        return {"ok": False, "records": [], "error": "Couldn't find the 17-column register table on this page. "
                "If it is sideways, rotate it upright first; if it isn't a Register Dakhil Kharij page, type the entry in by hand."}
    cols = grid["cols"]
    records = []
    for y0, y1 in grid["rows"]:
        def rect(col, top=0.0, bottom=1.0):
            x0, x1 = cols[col - 1]
            pad = (x1 - x0) * 0.04
            return (x0 + pad, y0 + (y1 - y0) * top + 0.003, x1 - pad, y0 + (y1 - y0) * bottom - 0.003)

        # Column 2 also holds an owner ID / share fraction lower down in some layouts; the khata stack
        # sits in the upper part of the cell, so read the top 60%.
        cells = {1: rect(1, 0, 0.5), 2: rect(2, 0, 0.6), 14: rect(14, 0, 0.5), 17: rect(17, 0, 0.5)}
        with lock:
            inked = {c: cell_ink(page, rot, r) >= INK_MIN for c, r in cells.items()}
            png = {c: cell_png(page, rot, r) for c, r in cells.items() if inked[c]}
        if not any(inked.values()):
            continue  # empty ruled row
        notes, flags, conf = [], [], "high"

        award = parse_award(ask(reader_model, PROMPT_PLAIN, png[1])) if inked[1] else ""
        if inked[1] and not award:
            notes.append("award number could not be read"); flags.append("award"); conf = "low"

        khata = ""
        if inked[2]:
            items = parse_khata(ask(khata_model, PROMPT_KHATA, png[2]))
            check = _plain_numbers(ask(reader_model, PROMPT_PLAIN, png[2]))
            pool, missing = list(check), []
            for n in (n for a, b in items for n in (a, b) if n is not None):
                if n in pool:
                    pool.remove(n)
                else:
                    missing.append(n)
            if missing or not items:
                # The two readings disagree: keep the digits both models saw, flag the row.
                conf = "medium"; flags.append("khata")
                notes.append(f"khata numbers uncertain (read {format_khata(items) or '-'}; check read {', '.join(map(str, check)) or '-'})")
                if not items:
                    items = [(n, None) for n in check]
            khata = format_khata(items)

        kanal = marla = sarsai = katat = ""
        if inked[14]:
            kms = parse_kms(ask(reader_model, PROMPT_PLAIN, png[14]))
            if kms:
                kanal, marla, sarsai = kms
            else:
                notes.append("column 14 area could not be read"); flags.append("area"); conf = "low"
        if inked[17]:
            total, katat = parse_total_katat(ask(reader_model, PROMPT_PLAIN, png[17]))
            if not katat:
                notes.append("katat under the column 17 total could not be read"); flags.append("katat"); conf = "low"
        continuation = (inked[1] or inked[2]) and not inked[14] and not inked[17]
        records.append({"award": award, "khata": khata, "katat": katat, "kanal": kanal, "marla": marla,
                        "sarsai": sarsai, "continuation": continuation, "confidence": conf, "notes": "; ".join(notes),
                        "flags": flags,
                        "cells": {str(c): r for c, r in cells.items()}})
    return {"ok": True, "records": records, "error": "", "grid": grid}


def header_segments(page, rot, table_top, min_gap=0.02):
    """The line of text just above the table holds 'محال …  حدبست نمبر …  تحصیل …  ضلع …  سال …'.
    Split it into phrases at wide blank gaps. Returns [(x0, x1, png)] in fractions, right to left."""
    z = RENDER_LONG / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), colorspace=pymupdf.csGRAY, alpha=False)
    w, h, D = _dark_rows(pix)
    y1 = int(table_top * h) - 3
    # Walk up from the table border to find the text line: the first run of inked rows.
    inked = [y for y in range(max(0, y1 - int(0.12 * h)), y1) if sum(D[y]) > 0.002 * w]
    if not inked:
        return []
    bottom = inked[-1]
    line_rows = [bottom]
    for y in range(bottom - 1, max(0, bottom - int(0.08 * h)), -1):
        if sum(D[y]) > 0.002 * w:
            line_rows.append(y)
        elif bottom - y > 0.012 * h and sum(D[y]) == 0:
            break
    ya, yb = min(line_rows) - 4, bottom + 4
    cols = [any(D[y][x] for y in range(ya, yb)) for x in range(w)]
    segs, start, gap = [], None, 0
    for x, ink in enumerate(cols + [False] * int(min_gap * w + 2)):
        if ink:
            if start is None:
                start = x
            gap = 0
        elif start is not None:
            gap += 1
            if gap > min_gap * w:
                segs.append((start, x - gap))
                start = None
    out = []
    for a, b in reversed(segs):  # Urdu reads right to left
        rect = ((a - 6) / w, ya / h, (b + 6) / w, yb / h)
        out.append((rect[0], rect[2], cell_png(page, rot, rect, longest=700)))
    return out


def _unrotate(rect, rot):
    """Fractions on the page as displayed (turned `rot`° clockwise) -> fractions of the unrotated page."""
    def pt(u, v):
        return {0: (u, v), 90: (v, 1 - u), 180: (1 - u, 1 - v), 270: (1 - v, u)}[rot]
    (ax, ay), (bx, by) = pt(rect[0], rect[1]), pt(rect[2], rect[3])
    return min(ax, bx), min(ay, by), max(ax, bx), max(ay, by)


def cell_ink(page, rot, frac_rect):
    """Share of dark pixels in a cell (borders trimmed). Blank cells are never sent to a model,
    because that is exactly where small models invent text."""
    x0, y0, x1, y1 = frac_rect
    dx, dy = (x1 - x0) * 0.06, (y1 - y0) * 0.02
    png = cell_png(page, rot, (x0 + dx, y0 + dy, x1 - dx, y1 - dy), longest=300)
    pix = pymupdf.Pixmap(png)
    if pix.n > 1:
        pix = pymupdf.Pixmap(pymupdf.csGRAY, pix)
    s = pix.samples
    return sum(1 for v in s if v < 140) / max(1, len(s))


def cell_png(page, rot, frac_rect, longest=900):
    """Render one cell (fractions of the rotated page) as PNG for the model."""
    x0, y0, x1, y1 = _unrotate(frac_rect, rot)
    r = page.rect
    clip = pymupdf.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height, r.x0 + x1 * r.width, r.y0 + y1 * r.height)
    zoom = max(1.0, min(longest / max(clip.width, clip.height), 8.0))
    return page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom).prerotate(rot), clip=clip, alpha=False).tobytes("png")
