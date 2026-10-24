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
import itertools
import math
import re

import pymupdf

RENDER_LONG = 2000  # px, longest side used for finding the grid
INK_MIN = 0.004     # a cell with less writing than this (see cell_ink) is treated as empty

PROMPT_PLAIN = ("Transcribe all text in this image exactly as written, top to bottom, one line per printed line. "
                "Keep digits, hyphens and slashes exactly. Output only the text.")

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


def parse_khasra(text):
    """Column 16 lists khasra numbers like '37/24/1', each with a short number under its bar ('37').
    Return the full khasra numbers in printed order, repeats included (each is a separate entry with
    its own area): '37/24/1, 37/23/2, 47/3/1, 37/23/2'. The under-bar numbers are dropped."""
    t = re.sub(r"[ \t]*/[ \t]*", "/", _norm(text))  # '56 / 7 / 2' -> '56/7/2', without joining lines
    return ", ".join(re.findall(r"(?<![\d/])\d{1,5}(?:/\d{1,5})+(?![\d/])", t))


def kms_units(k, m, s, per):
    return float(k) * 20 * per + float(m) * per + float(s)


def pair_khasra_areas(khasras, areas):
    """Column 16 lists khasra numbers and column 17 the area of each, line for line; column 17 may end
    with the total of those areas. Returns (items, total, problem): items = [{"khasra", "area"}] with
    area as "K-M-S" ("" if missing), total = "K-M-S" or "", problem = a note when the counts or the
    sum don't line up."""
    fmt = lambda a: "-".join(a)
    total, problem = "", ""
    if areas and len(areas) == len(khasras) + 1:
        *areas, last = areas
        total = fmt(last)
        # The third figure may be Sarsai (9 to a Marla) or sq ft (272): accept either.
        if not any(abs(sum(kms_units(*a, per) for a in areas) - kms_units(*last, per)) < 0.01 for per in (9, 272)):
            problem = f"column 17 areas don't add up to the total printed under them ({total})"
    elif areas and len(areas) != len(khasras):
        problem = f"column 16 lists {len(khasras)} khasra number(s) but column 17 has {len(areas)} area(s) — check which area goes with which khasra"
    items = [{"khasra": k, "area": fmt(areas[i]) if i < len(areas) else ""} for i, k in enumerate(khasras)]
    return items, total, problem


def _plain_numbers(text):
    """Whole numbers in a transcription, skipping fractions (26/1601), CNIC-style IDs and K-M-S areas."""
    t = KMS.sub(" ", _norm(text))
    t = re.sub(r"\d+\s*/\s*\d+", " ", t)
    t = re.sub(r"\d+(?:-\d+){1,}", " ", t)
    return [int(n) for n in re.findall(r"(?<!\d)(\d{1,6})(?!\d)", t)]


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


def estimate_skew(page, rot=0):
    """Small scan tilt in degrees (positive = content slopes down to the right). Projection-profile
    method: rows of a straight table/text give the 'peakiest' row histogram; try angles, keep the best."""
    z = 1000 / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), colorspace=pymupdf.csGRAY, alpha=False)
    w, h, s, st = pix.width, pix.height, pix.samples, pix.stride
    pts = [(x, y) for y in range(h) for x in range(w) if s[y * st + x] < 160]
    if len(pts) < 500:
        return 0.0
    if len(pts) > 60000:
        pts = pts[:: len(pts) // 60000 + 1]

    def score(deg):
        t = math.tan(math.radians(deg))
        hist = {}
        for x, y in pts:
            k = int(y - x * t)
            hist[k] = hist.get(k, 0) + 1
        return sum(v * v for v in hist.values())

    best = max((score(a / 2), a / 2) for a in range(-6, 7))[1]                 # coarse: 0.5° steps
    best = max((score(best + a / 20), best + a / 20) for a in range(-10, 11))[1]  # fine: 0.05° steps
    return best


def straightened(page, rot, skew, longest=3000):
    """Render the page upright and de-tilted, and return it as a one-page document to read from."""
    z = longest / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot - skew), alpha=False)
    img = pymupdf.open("png", pix.tobytes("png"))
    return pymupdf.open("pdf", img.convert_to_pdf())[0]


# Column rules of the printed Register Dakhil Kharij, left to right, as fractions of the table width
# (measured on a clean page). Only used to choose the right 18 when a scan shows extra rules.
FORM_RULES = [0.0, 0.058, 0.1235, 0.1927, 0.2411, 0.2916, 0.3491, 0.4635, 0.5051, 0.5567,
              0.6094, 0.6631, 0.71, 0.7616, 0.818, 0.9266, 0.9611, 1.0]


def pick_columns(xs):
    """From candidate vertical rules, return the 18 that form the 17 columns, or None."""
    if len(xs) == 18:
        return xs
    if not 18 < len(xs) <= 22:
        return None
    best, best_err = None, 0.04  # worst allowed deviation from the form layout, in table widths
    for drop in itertools.combinations(range(len(xs)), len(xs) - 18):
        cand = [x for i, x in enumerate(xs) if i not in drop]
        span = cand[-1] - cand[0]
        if span <= 0:
            continue
        err = max(abs((x - cand[0]) / span - t) for x, t in zip(cand, FORM_RULES))
        if err < best_err:
            best, best_err = cand, err
    return best


def find_grid(page, rot=0):
    """Locate the register table on an (upright after `rot`) page.
    Returns {"cols": [(x0, x1) for column 1..17], "rows": [(y0, y1) data bands], "top": y, "size": (w, h)}
    in fractions of the rotated page, or None if no 17-column ruled table is found."""
    z = RENDER_LONG / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), colorspace=pymupdf.csGRAY, alpha=False)
    w, h, D = _dark_rows(pix)

    # Horizontal rules: rows that are largely dark. OR-ing a few rows tolerates slight skew; the bar is
    # deliberately low because scanned borders are often faint or broken (a real scan's top border
    # covered only ~45% of the row).
    counts = []
    for y in range(h - 4):
        rows5 = D[y:y + 5]
        counts.append(sum(1 for i in range(w) if any(r[i] for r in rows5)))
    # Rules are thin: a dense row of text can also cover much of the width, but over many more rows.
    runs, cur = [], []
    for y, n in enumerate(counts):
        if n > 0.45 * w and (not cur or y - cur[-1] <= 2):
            cur.append(y)
        elif n > 0.45 * w:
            runs.append(cur); cur = [y]
    if cur:
        runs.append(cur)
    lines = [(r[0] + r[-1]) // 2 + 2 for r in runs if len(r) <= 14]
    if len(lines) < 3:
        return None

    def verticals(y0, y1, frac=0.9):
        y0, y1 = y0 + 3, y1 - 3
        if y1 - y0 < 6:
            return []
        need = frac * (y1 - y0)
        edge = int(0.01 * w)  # the scanned paper's own edge often shows up as a dark "rule"
        xs = [x for x in range(edge, w - 2 - edge)
              if sum(1 for y in range(y0, y1) if D[y][x] or D[y][x + 1] or D[y][x + 2]) >= need]
        return _group(xs, gap=8)  # thick or doubled rules count once

    bands = [(a, b, verticals(a, b, 0.8)) for a, b in zip(lines, lines[1:])]
    # The column-number row (١ … ١٧): the first SHORT band crossed by 18 rules (17 columns). Short bands
    # are sensitive to how strictly a rule must span them, so try several strictness levels; extra
    # rules (a frame, the next page peeking in) are weeded out by matching the form's column layout.
    num = v = None
    for i, (a, b, _) in enumerate(bands):
        if b - a > 0.06 * h:
            continue
        for frac in (0.9, 0.8, 0.7, 0.6, 0.5):
            v = pick_columns(verticals(a, b, frac))
            if v:
                break
        if v:
            num = i
            break
    if num is None:
        return None
    cols =[(v[17 - k] / w, v[18 - k] / w) for k in range(1, 18)]  # column 1 is rightmost

    # Header: after the number row comes a row of merged group headings (few rules) and then the
    # column titles. Data rows are the ruled bands after that which still carry most column rules.
    i = num + 1
    if i < len(bands) and len(bands[i][2]) < 12:
        i += 1          # group-heading row
    i += 1              # column-title row
    # Everything below the titles, down to the table's bottom rule, is entry rows (their rules may be
    # faint on continuation pages, so don't insist on seeing every column rule).
    rows = [(a / h, b / h) for a, b, vv in bands[i:] if (b - a) > 0.02 * h]
    return {"cols": cols, "rows": rows, "top": bands[num][0] / h, "size": (w, h)}


def _gray(png):
    pix = pymupdf.Pixmap(png)
    if pix.n > 1:
        pix = pymupdf.Pixmap(pymupdf.csGRAY, pix)
    return pix.width, pix.height, pix.samples, pix.stride


def khata_layout(page, rot, rect, min_gap=3):
    """The printed lines of a column-2 cell, top to bottom, as (kind, sub_rect): kind is "num" for a
    clearly number-shaped line (3+ digits are 2-3x wider than tall, even the small ones under a
    fraction bar) or "?" for a squat line, which is either the word تا or a short number — the caller
    asks the model which. Fraction bars, specks and slivers of ruled lines are dropped."""
    x0, y0, x1, y1 = rect
    dx = (x1 - x0) * 0.08  # stay clear of the column rules
    inner = (x0 + dx, y0, x1 - dx, y1)
    w, h, s, st = _gray(cell_png(page, rot, inner, longest=700))
    dark = lambda x, y: s[y * st + x] < 150

    def longest_run(y):
        best = cur = 0
        for x in range(w):
            cur = cur + 1 if dark(x, y) else 0
            best = max(best, cur)
        return best

    # A fraction bar is a long solid stroke; digits never are. Treat bar rows as blank so a number
    # pressed right under its bar still counts as its own line.
    bar_rows = {y for y in range(h) if longest_run(y) >= 0.22 * w}
    ink = [0 if y in bar_rows else sum(1 for x in range(w) if dark(x, y)) for y in range(h)]
    bands, start, gap = [], None, 0
    for y, n in enumerate(ink + [0] * (min_gap + 1)):
        if n > 0:
            start = y if start is None else start
            gap = 0
        elif start is not None:
            gap += 1
            if gap > min_gap:
                bands.append((start, y - gap))
                start = None
    out = []
    for a, b in bands:
        xs = [x for x in range(w) if any(dark(x, y) for y in range(a, b + 1) if y not in bar_rows)]
        if not xs:
            continue
        height, width = b - a + 1, xs[-1] - xs[0] + 1
        if width >= 6 * height or (height <= 3 and width <= 4) or width < 0.4 * height or width <= 3:
            continue  # fraction bar, speck, or a sliver of a ruled line
        pad = max(2, height // 4)
        sub = (xs[0] - pad, max(0, a - pad), xs[-1] + pad, min(h, b + pad))
        sub = (inner[0] + (inner[2] - inner[0]) * sub[0] / w, y0 + (y1 - y0) * sub[1] / h,
               inner[0] + (inner[2] - inner[0]) * sub[2] / w, y0 + (y1 - y0) * sub[3] / h)
        out.append(("num" if width >= 1.9 * height else "?", sub))
    return out


def build_khata(numbers, layout):
    """Put the digits read from the whole cell onto the line layout: a "to" line between two numbers
    makes a range. Returns (items, ok) — ok is False when the counts don't line up."""
    kinds = [k for k, *_ in layout]
    slots = kinds.count("num")
    numbers = list(numbers)
    while len(numbers) > slots and any(n < 10 for n in numbers):
        # the model sometimes reads تا as a single digit ("5"): drop such strays, keeping the order
        numbers.remove(next(n for n in numbers if n < 10))
    if len(numbers) != slots:
        return [(n, None) for n in numbers], False
    items, it, pending = [], iter(numbers), False
    for k in kinds:
        if k == "to":
            pending = bool(items) and items[-1][1] is None
        elif pending:
            items[-1], pending = (items[-1][0], next(it)), False
        else:
            items.append((next(it), None))
    return items, True


def read_page(page, rot, ask, reader_model, urdu_model, lock=None, plot_field="katat"):
    """plot_field: "katat" = the small number under the column 17 total (some registers);
    "khasra" = the khasra numbers listed in column 16 (registers that list khasra numbers instead).
    The value is returned in the record's "katat" key either way, with the flag "katat"."""
    """Read every mutation entry on one register page.
    `ask(model, prompt, png) -> str` calls the local vision model. `lock` (optional) is held while
    rendering — MuPDF isn't thread-safe — but not during the slow model calls. Returns a dict:
    {"ok": bool, "error": str, "records": [{award, khata, katat, kanal, marla, sarsai, continuation,
     confidence, notes, cells}], "grid": {...}} — cells/grid (fractions) let the UI outline what was read."""
    lock = lock or contextlib.nullcontext()
    with lock:
        skew = estimate_skew(page, rot)
        if abs(skew) >= 0.1:  # tilted scan: read from an upright, de-tilted copy
            page, rot = straightened(page, rot, skew, longest=3600), 0
        grid = find_grid(page, rot)
    if not grid:
        return {"ok": False, "records": [], "error": "Couldn't find the 17-column register table on this page. "
                "If it is sideways, rotate it upright first; if it isn't a Register Dakhil Kharij page, type the entry in by hand."}
    cols = grid["cols"]
    with lock:
        location = header_box(page, rot, grid)
    records = []
    for y0, y1 in grid["rows"]:
        def rect(col, top=0.0, bottom=1.0):
            x0, x1 = cols[col - 1]
            pad = (x1 - x0) * 0.04
            return (x0 + pad, y0 + (y1 - y0) * top + 0.003, x1 - pad, y0 + (y1 - y0) * bottom - 0.003)

        # The khata stack can run most of the way down column 2; the other values sit near the top.
        cells = {1: rect(1, 0, 0.5), 2: rect(2, 0, 1.0), 14: rect(14, 0, 0.5), 17: rect(17, 0, 0.5)}
        if plot_field == "khasra":
            # A list of khasra numbers can run down the cell, and long ones touch the right-hand rule,
            # so take the full height and reach a little past that rule.
            x0, x1 = cols[15]
            cells[16] = (x0 + (x1 - x0) * 0.04, y0 + 0.003, x1 + (x1 - x0) * 0.06, y1 - 0.003)
            cells[17] = rect(17, 0, 1.0)  # one area per khasra, line for line, sometimes a total at the end
        with lock:
            inked = {c: cell_ink(page, rot, r) >= INK_MIN for c, r in cells.items()}
            png = {c: cell_png(page, rot, r) for c, r in cells.items() if inked[c]}
        if not (inked[1] or inked[2]):
            continue  # every entry has an award number / khata; anything else is an empty or stray row
        notes, flags, conf = [], [], "high"

        award = parse_award(ask(reader_model, PROMPT_PLAIN, png[1])) if inked[1] else ""
        if inked[1] and not award:
            notes.append("award number could not be read"); flags.append("award"); conf = "low"

        khata = ""
        if inked[2]:
            # Digits from one whole-cell reading (accurate); where the ranges are from the line layout.
            numbers = _plain_numbers(ask(reader_model, PROMPT_PLAIN, png[2]))
            with lock:
                layout = khata_layout(page, rot, cells[2])
                unsure = {i: cell_png(page, rot, r, longest=200) for i, (k, r) in enumerate(layout) if k == "?"}
            for i, img in unsure.items():
                # Squat line: the word تا, or a short number? TranslateGemma transcribes a short number
                # and returns nothing for تا; qwen tends to invent digits for تا, so it isn't used here.
                digits = re.search(r"\d", _norm(ask(urdu_model, PROMPT_PLAIN, img)))
                layout[i] = ("num" if digits else "to", layout[i][1])
            items, ok = build_khata(numbers, layout)
            if not ok or not items:
                conf = "medium"; flags.append("khata")
                notes.append(f"khata: read {len(numbers)} number(s) but the cell shows "
                             f"{sum(k == 'num' for k, *_ in layout)} line(s) of numbers — check the list and any 'to' ranges")
            khata = format_khata(items)

        kanal = marla = sarsai = katat = ""
        if inked[14]:
            kms = parse_kms(ask(reader_model, PROMPT_PLAIN, png[14]))
            if kms:
                kanal, marla, sarsai = kms
            else:
                notes.append("column 14 area could not be read"); flags.append("area"); conf = "low"
        khasra_items, col17_total = [], ""
        if plot_field == "khasra":
            if inked[16]:
                katat = parse_khasra(ask(reader_model, PROMPT_PLAIN, png[16]))
                if not katat:
                    notes.append("khasra number in column 16 could not be read"); flags.append("katat"); conf = "low"
            areas = KMS.findall(_norm(ask(reader_model, PROMPT_PLAIN, png[17]))) if inked[17] else []
            khasra_items, col17_total, msg = pair_khasra_areas(katat.split(", ") if katat else [], areas)
            if msg:
                notes.append(msg); flags.append("khasra_area"); conf = "medium"
        elif inked[17]:
            total, katat = parse_total_katat(ask(reader_model, PROMPT_PLAIN, png[17]))
            if not katat:
                notes.append("katat under the column 17 total could not be read"); flags.append("katat"); conf = "low"
        continuation = not inked[14] and not inked[17]
        khasra_only = False
        if not khata and not (kanal or marla or sarsai):
            # Only more owners for the award, or a summary page (totals, share-transferred box) with no
            # khata and no column-14 area: nothing to put in the table — unless it lists khasra numbers
            # (khasra mode), which the caller adds if they aren't in the table yet.
            if not (plot_field == "khasra" and katat):
                continue
            khasra_only = True  # keep only the khasra/area notes; khata and area are blank by design here
            keep = [i for i, f in enumerate(flags) if f in ("katat", "khasra_area")]
            notes = [n for n in notes if n.startswith(("khasra", "column 16", "column 17"))]
            flags = [flags[i] for i in keep]
            conf = "medium" if flags else "high"
        records.append({"award": award, "khata": khata, "katat": katat, "kanal": kanal, "marla": marla,
                        "sarsai": sarsai, "continuation": continuation, "confidence": conf, "notes": "; ".join(notes),
                        "flags": flags, "khasra_only": khasra_only,
                        "khasra_items": khasra_items, "col17_total": col17_total,
                        "cells": {str(c): r for c, r in cells.items()}})  # the cells read, for outlining
    return {"ok": True, "records": records, "error": "", "grid": grid, "skew": round(skew, 2), "location": location}


def header_box(page, rot, grid):
    """Box around the line just above the table that names the place — 'محال …  حدبست نمبر …
    تحصیل …  ضلع …'. The names aren't read (the local models can't read this handwriting reliably);
    the box lets the user check the batch District/Tehsil/Mouza against the page at a glance."""
    z = RENDER_LONG / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), colorspace=pymupdf.csGRAY, alpha=False)
    w, h, D = _dark_rows(pix)
    xa, xb = int(grid["cols"][16][0] * w), int(grid["cols"][0][1] * w)  # the table's left and right edges
    top = int(grid["top"] * h) - 4
    inked = lambda y: sum(D[y][xa:xb]) > 0.003 * (xb - xa)
    # Walk up from the table's top border to the nearest line of text, then to its top.
    y = top
    while y > top - 0.08 * h and not inked(y):
        y -= 1
    if not inked(y):
        return None
    bottom = y
    while y > max(0, bottom - 0.06 * h) and (inked(y) or inked(y - 1) or inked(y - 2)):
        y -= 1
    pad = 0.004
    return (xa / w, max(0.0, y / h - pad), xb / w, min(grid["top"], bottom / h + pad))


def _unrotate(rect, rot):
    """Fractions on the page as displayed (turned `rot`° clockwise) -> fractions of the unrotated page."""
    def pt(u, v):
        return {0: (u, v), 90: (v, 1 - u), 180: (1 - u, 1 - v), 270: (1 - v, u)}[rot]
    (ax, ay), (bx, by) = pt(rect[0], rect[1]), pt(rect[2], rect[3])
    return min(ax, bx), min(ay, by), max(ax, bx), max(ay, by)


def cell_ink(page, rot, frac_rect):
    """Amount of writing in a cell (borders trimmed), as dark pixels per square of the cell's width —
    so one line of text counts the same in a short cell and in a tall, mostly blank one. Blank cells
    are never sent to a model, because that is exactly where small models invent text."""
    x0, y0, x1, y1 = frac_rect
    dx, dy = (x1 - x0) * 0.06, (y1 - y0) * 0.02
    aspect = ((y1 - y0) * page.rect.height) / max(1e-6, (x1 - x0) * page.rect.width)
    longest = int(150 * max(1.0, aspect))  # cell width ≈ 150 px whatever its height
    w, h, s, st = _gray(cell_png(page, rot, (x0 + dx, y0 + dy, x1 - dx, y1 - dy), longest=longest))
    dark = [[s[y * st + x] < 140 for x in range(w)] for y in range(h)]
    # Ignore ruled lines that stray into the crop (a border that's still slightly sloped after
    # straightening): pixels in long straight runs are rules, not writing.
    rule = [[False] * w for _ in range(h)]
    for y in range(h):
        x = 0
        while x < w:
            if dark[y][x]:
                e = x
                while e < w and dark[y][e]:
                    e += 1
                if e - x >= 0.25 * w:
                    for i in range(x, e):
                        rule[y][i] = True
                x = e
            else:
                x += 1
    for x in range(w):
        y = 0
        while y < h:
            if dark[y][x]:
                e = y
                while e < h and dark[e][x]:
                    e += 1
                if e - y >= 0.25 * h:
                    for i in range(y, e):
                        rule[i][x] = True
                y = e
            else:
                y += 1
    # A rule that is very slightly slanted breaks into short straight pieces, but it is still one
    # continuous stroke if a pixel of wobble either side is allowed. Writing has gaps (between
    # characters, between lines), so it never forms such a long run.
    for x in range(w):
        lo, hi = max(0, x - 1), min(w, x + 2)
        y = 0
        while y < h:
            if any(dark[y][lo:hi]):
                e = y
                while e < h and any(dark[e][lo:hi]):
                    e += 1
                if e - y >= 0.3 * h:
                    for i in range(y, e):
                        for xx in range(lo, hi):
                            rule[i][xx] = True
                y = e
            else:
                y += 1
    for y in range(h):
        rows3 = dark[max(0, y - 1):y + 2]
        x = 0
        while x < w:
            if any(r[x] for r in rows3):
                e = x
                while e < w and any(r[e] for r in rows3):
                    e += 1
                if e - x >= 0.5 * w:
                    for yy in range(max(0, y - 1), min(h, y + 2)):
                        for i in range(x, e):
                            rule[yy][i] = True
                x = e
            else:
                x += 1
    # Column rules sit at the cell's edges; after straightening a scan they can still lean across a
    # few percent of a tall cell's width, in broken pieces. Writing has plenty of ink further in.
    edge = int(0.1 * w)
    n = sum(1 for y in range(h) for x in range(edge, w - edge) if dark[y][x] and not rule[y][x])
    return n / max(1, w * w)


def cell_png(page, rot, frac_rect, longest=900):
    """Render one cell (fractions of the rotated page) as PNG for the model."""
    x0, y0, x1, y1 = _unrotate(frac_rect, rot)
    r = page.rect
    clip = pymupdf.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height, r.x0 + x1 * r.width, r.y0 + y1 * r.height)
    zoom = max(1.0, min(longest / max(clip.width, clip.height), 8.0))
    return page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom).prerotate(rot), clip=clip, alpha=False).tobytes("png")
