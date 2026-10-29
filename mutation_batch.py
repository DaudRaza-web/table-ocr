"""Batch mutation extraction: every mutation PDF under a root -> an Excel table next to it.

Run from the web UI's "Mutation batch" tab (server.py), through run().

Each mutation document (Register Dakhil Kharij pages) is read with the register reader (mutation.py,
khasra mode). One row per khasra: Mutation No, Khata Khatuni, Khasra No.; the areas are not listed
per khasra, only the document's total at the end: its grand total (میزان کل) when printed, otherwise
the khatas' printed sub-totals added up - one total for each mutation number in the document. District / Tehsil / Mouza come from the award document in the
same folder (its converted Excel file if there is one), not from the mutation. A Word file with the
same name lists the place, then per mutation each khata with its khasras and the mutation's total area.

Documents that can't be read, or with more than 30% of their data missing, get a
"<name> - REVIEW MANUALLY.txt" note next to them. The models are checked every few pages and after any
page that fails, and reloaded / restarted when they have degraded.
"""
import datetime
import os
import re
import time
import traceback

import pymupdf

import award
import award_batch
import mutation

HEADERS = ["S.No.", "Mutation No", "District", "Tehsil", "Mouza", "Khata Khatuni", "Khasra No.", "Kanal", "Marla",
           "Sarsai / Sq ft", "Page"]
WIDTHS = (7, 12, 16, 16, 18, 22, 14, 8, 8, 12, 7)
HEALTH_EVERY = 10    # pages between model checks
MISSING_LIMIT = 0.30  # more than this share of the data missing -> review by hand


def is_mutation(name):
    n = name.lower()
    return n.endswith(".pdf") and ("mutation" in n or "intiqal" in n) and "correction" not in n


def _stem(path):
    return re.sub(r"\.pdf$", "", os.path.basename(path), flags=re.I)


def output_name(path, numbers, place):
    nos = " & ".join(numbers) or _stem(path)
    mouza = re.sub(r'[<>:"/\\|?*]', "-", place.get("mouza", "")).strip()
    # a file whose name doesn't give its mutation numbers (a phone scan) is named in brackets, so a later
    # run can tell it has been done
    src = f"({_stem(path)})" if numbers and not numbers_in_name(path) else ""
    return " ".join(x for x in ("Mutation", nos, src, mouza, "- Land.xlsx") if x)


def review_name(path):
    return re.sub(r"\.pdf$", "", path, flags=re.I) + " - REVIEW MANUALLY.txt"


def numbers_in_name(path):
    """"Mutation Nos. 3272 & 3273 Village ..." -> ["3272", "3273"]"""
    m = re.search(r"mutation\s*nos?\.?\s*([\d\s&,and]+)", os.path.basename(path), re.I)
    return re.findall(r"\d+", m.group(1)) if m else []


def is_unnamed(name):
    """A PDF whose name says neither mutation nor award (e.g. "CamScanner 29-10-2026 15.49.pdf"): what it
    is has to be checked from its pages (classify)."""
    low = name.lower()
    return low.endswith(".pdf") and not is_mutation(name) and "award" not in low and "correction" not in low


def find_documents(root):
    """-> [(folder, [mutation pdfs by name], [pdfs to check by content])] for every folder under root that
    holds either."""
    out = []
    for folder, dirs, names in os.walk(root):
        dirs.sort()
        files = [os.path.join(folder, n) for n in sorted(names) if is_mutation(n)]
        unnamed = [os.path.join(folder, n) for n in sorted(names) if is_unnamed(n)]
        if files or unnamed:
            out.append((folder, files, unnamed))
    return out


def is_done(path, names):
    """A document is done when its Excel file or its review note exists."""
    stem = _stem(path)
    if os.path.basename(review_name(path)) in names:
        return True
    if not numbers_in_name(path):
        return any(n.startswith("Mutation") and n.endswith("- Land.xlsx") and f"({stem})" in n for n in names)
    return any(n.startswith("Mutation") and n.endswith("- Land.xlsx") and _same_doc(n, stem) for n in names)


def _same_doc(xlsx, stem):
    nums = numbers_in_name(stem + ".pdf")
    return bool(nums) and all(re.search(rf"\b{n}\b", xlsx) for n in nums)


# ------------------------------------------------------------------- place --

def place_from_award(folder, root, ask_json, log, found_awards=()):
    """District / Tehsil / Mouza from an award in this folder (or the folders above it, up to root):
    its converted Excel file if there is one, else read from the award PDF (named so, or found by
    content: `found_awards`). -> (place, source)"""
    root = os.path.normpath(root)
    here = os.path.normpath(folder)
    while True:
        names = sorted(os.listdir(here))
        for n in names:  # an award already converted: no model needed
            if award_batch.is_output(n):
                try:
                    from openpyxl import load_workbook
                    ws = load_workbook(os.path.join(here, n), read_only=True)["Award Land"]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row and any(row[2:5]):
                            return ({"district": row[2] or "", "tehsil": row[3] or "", "mouza": row[4] or ""},
                                    os.path.join(here, n))
                except Exception:
                    pass
        awards = [os.path.join(here, n) for n in names if award_batch.is_award(n)] or \
            [a for a in found_awards if os.path.normpath(os.path.dirname(a)) == here]
        if awards:
            src = max(awards, key=award_batch.award_key)
            log(f"  reading District/Tehsil/Mouza from {os.path.basename(src)}")
            doc = pymupdf.open(src)
            place = {}
            for n in range(min(doc.page_count, 6)):  # the land table's first page
                if award.find_table(award._gray(doc[n], dpi=200)):
                    place = award.read_place(doc[n], ask_json)
                    break
            place, _ = award.fill_place(place, award.read_place_doc(doc[0], ask_json))
            return place, src
        if here == root or os.path.dirname(here) == here or not here.startswith(root):
            return {}, ""
        here = os.path.dirname(here)


# ---------------------------------------------------------------- document --

def _kms(text):
    m = mutation.KMS.search(text or "")
    return (int(m.group(1)), int(m.group(2)), float(m.group(3))) if m else None


def _fmt3(v):
    return str(int(v)) if float(v).is_integer() else str(v)


def add_areas(areas):
    """Sum of K-M-S areas. The third figure is Sarsai (9 to a Marla) or Sq ft (272 to a Marla); a third
    figure of 9 or more means Sq ft. -> ((K, M, S), unit)"""
    per, unit = (272, "Sq ft") if any(a[2] >= 9 for a in areas) else (9, "Sarsai")
    units = sum(a[0] * 20 * per + a[1] * per + a[2] for a in areas)
    k = int(units // (20 * per))
    units -= k * 20 * per
    m = int(units // per)
    return (k, m, round(units - m * per, 2)), unit


def read_document(path, srv, ask, models, place, log, live=None, stop=lambda: False, ask_json=None):
    """One mutation PDF -> {"rows", "totals" (one per mutation number), "issues", "missing", "unreadable",
    "numbers", "pages"}. live(rows, page, pages) is called after each page with the table so far.
    Each page is read cell by cell from its ruled grid; a page whose grid can't be found either way up is
    read whole by the vision model (`ask_json`), and flagged."""
    doc = pymupdf.open(path)
    named = numbers_in_name(path)
    rows, subtotals, issues, totals_only = [], [], [], []
    expected = missing = failed = 0
    last_mut, last_khata = (named[0] if len(named) == 1 else ""), {}
    since_check = 0
    footer_misses = 0  # pages whose footer had no total: after two, footers aren't read for this document
    reader, urdu = models

    def table():
        return [[i, r["mut"], place.get("district", ""), place.get("tehsil", ""), place.get("mouza", ""),
                 r["khata"], r["khasra"], "", "", "", r["page"]] for i, r in enumerate(rows, 1)]

    for n in range(doc.page_count):
        if stop():
            issues.append(f"Stopped by the user at page {n + 1} of {doc.page_count}")
            break
        log(f"    page {n + 1}/{doc.page_count}")
        if since_check >= HEALTH_EVERY:
            heal(srv, models, log)
            since_check = 0
        rot = srv.detect_orientation(doc, n)[0]  # register pages are landscape forms, often scanned on their side
        res = None
        for attempt in range(2):
            try:
                res = mutation.read_page(doc[n], rot, ask, reader, urdu, lock=getattr(srv, "DOCS_LOCK", None),
                                         plot_field="khasra", keep_totals=True, read_footer=footer_misses < 2)
            except Exception as e:
                res = {"ok": False, "records": [], "error": f"{type(e).__name__}: {e}", "crashed": True}
            bad = res.get("crashed") or (res["ok"] and any(
                f in ("award", "katat") for rec in res["records"] for f in rec.get("flags", [])))
            if not res["ok"] and not res.get("crashed"):
                # No register grid: the scans in one document don't all face the same way, so try the
                # page turned the other way round before giving up on it.
                other = mutation.read_page(doc[n], (rot + 180) % 360, ask, reader, urdu,
                                           lock=getattr(srv, "DOCS_LOCK", None), plot_field="khasra", keep_totals=True,
                                           read_footer=footer_misses < 2)
                if other["ok"]:
                    res = other
                elif ask_json:
                    v = vision_page(doc[n], rot, ask_json)
                    if v["ok"] and v["records"]:
                        res = v
                        issues.append(f"Page {n + 1}: no ruled grid found - read from the whole page instead; check "
                                      "its rows against the scan")
                        log(f"    page {n + 1}: no grid - read from the whole page")
                    else:
                        res = dict(res, error=f"no register grid found, and {v['error'] or 'nothing could be read from the whole page'}")
                break
            if not bad or attempt:
                break
            h = heal(srv, models, log)  # a page that failed or came back unreadable: are the models fine?
            since_check = 0
            if h["action"] == "none":
                break  # the models are fine: the page is just hard to read
        since_check += 1
        if not res["ok"]:
            failed += 1
            issues.append(f"Page {n + 1}: {res['error']}")
            log(f"    !! page {n + 1} skipped: {'model error' if res.get('crashed') else 'no register table found (both turns, several tilts tried)'}")
            continue
        for rec in res["records"]:
            mut = rec["award"] or last_mut
            if rec["award"]:
                last_mut = rec["award"]
            if rec.get("whole_khata") is not None and not rec["khasra_items"]:
                # the whole khata changes hands and no khasra numbers are listed: one row says so, and its
                # area is the khata's sub-total
                plots = rec["whole_khata"]
                label = "Whole khata" + (f" ({plots} plot{'s' if plots != 1 else ''})" if plots else "")
                rec = dict(rec, khasra_items=[{"khasra": label, "area": ""}])
            if rec.get("total_only") or (not rec["khasra_items"] and rec["col17_total"] and not rec["khasra_only"]):
                # a total with no khasras: maybe this mutation's grand total
                totals_only.append((mut, rec["col17_total"], n + 1))
                continue
            if rec["continuation"] and not rec["khasra_items"]:
                continue  # more owners of the same entry
            if not rec["khasra_items"]:
                # khasra column unreadable (or blank) on an entry that should list them: its three
                # values (mutation no, khata, khasra) count towards the missing share
                expected += 3
                missing += 1 + (not mut) + (not rec["khata"])
                issues.append(f"Page {n + 1}: no khasra numbers read for khata {rec['khata'] or '?'} - check")
                continue
            khata = rec["khata"] or (last_khata.get(mut, "") if rec["khasra_only"] else "")
            if rec["khata"]:
                last_khata[mut] = rec["khata"]
            # A later page can repeat khasras already listed for this mutation and khata (a continuation
            # page, or the entry printed again with the order text): those are left out. The same khasra
            # under another khata is a separate entry and is kept, as are repeats within one entry.
            seen = {r["khasra"] for r in rows if r["mut"] == mut and r["khata"] == khata}
            new = [it for it in rec["khasra_items"] if it["khasra"] not in seen]
            for it in new:
                rows.append({"mut": mut, "khata": khata, "khasra": it["khasra"], "page": n + 1})
                expected += 3
                missing += (not mut) + (not khata)
            if rec["col17_total"] and (new or rec["khasra_only"]) and \
                    not any(s[:3] == (mut, khata, rec["col17_total"]) for s in subtotals):
                subtotals.append((mut, khata, rec["col17_total"], n + 1))
            for note in filter(None, (rec.get("notes") or "").split("; ")):
                if not note.startswith("column 14"):
                    issues.append(f"Page {n + 1}: {note}")
        if res.get("footer_total"):
            # the total printed under the table (computer-printed form): this mutation's grand total
            totals_only.append((last_mut, res["footer_total"], n + 1, "footer"))
            footer_misses = -10 ** 6  # this document prints footer totals: keep reading them
        elif res.get("footer_total") == "":
            footer_misses += 1
        if live:
            live(table(), n + 1, doc.page_count)

    numbers = list(dict.fromkeys(r["mut"] for r in rows if r["mut"])) or named
    totals = [mutation_total(m, subtotals, totals_only, issues) for m in (numbers or [""])]
    numbers = numbers or named
    ratio = (missing / expected) if expected else 1.0
    page_ratio = failed / doc.page_count if doc.page_count else 1.0
    unreadable = not rows
    return {"rows": rows, "table": table(), "totals": totals, "issues": issues, "missing": max(ratio, page_ratio),
            "unreadable": unreadable, "failed_pages": failed, "numbers": numbers, "pages": doc.page_count}


def mutation_total(mut, subtotals, totals_only, issues):
    """One mutation's total: its grand total (a column 17 area with no khasras after its entries, at
    least its khata sub-totals added up), else its printed sub-totals added up.
    -> {"mut", "total": (K, M, S) or None, "kind", "unit", "subs": [(mut, khata, area, page)]}"""
    subs = [s for s in subtotals if s[0] == mut]
    sub = [a for a in (_kms(s[2]) for s in subs) if a]
    sub_sum, unit = add_areas(sub) if sub else (None, "")
    label = f"Mutation {mut}" if mut else "The document"
    footers = [t for t in totals_only if t[0] == mut and t[3:] == ("footer",) and _kms(t[1])]
    if footers:
        # The total printed under the table ("میزان کل رقبہ منتقلہ", the area transferred) is the mutation's
        # total even when smaller than its khatas' areas. It repeats on every page: the most common reading
        # wins over a stray misread (or a table cell caught in the footer crop).
        votes = {}
        for t in footers:
            votes.setdefault(_kms(t[1]), []).append(t[2])
        g, pages = max(votes.items(), key=lambda kv: (len(kv[1]), -kv[1][0]))
        where = f"page {pages[0]}" if len(footers) == 1 else f"{len(pages)} of {len(footers)} pages, first page {pages[0]}"
        if len(votes) > 1:
            issues.append(f"{label}: the total under the table reads differently on some pages - "
                          f"{fmt_kms(g)} taken (most pages); check")
        return {"mut": mut, "total": g, "kind": f"total printed under the table ({where})", "subs": subs,
                "unit": "Sq ft" if g[2] >= 9 else "Sarsai"}
    grands = [t for t in totals_only if t[0] == mut]
    if grands:
        _m, area, page = grands[-1][:3]
        g = _kms(area)
        per = 272 if unit == "Sq ft" or (g and g[2] >= 9) else 9
        size = lambda a: a[0] * 20 * per + a[1] * per + a[2]  # noqa: E731
        if g and (not sub or size(g) + 0.01 >= size(sub_sum)):
            return {"mut": mut, "total": g, "kind": f"grand total printed on page {page}", "subs": subs,
                    "unit": unit or ("Sq ft" if per == 272 else "Sarsai")}
        if g:
            issues.append(f"{label}: area {area} on page {page} with no khasras is smaller than its khata sub-totals "
                          "added up, so it was not taken as the grand total")
    if sub:
        if len(sub) < len(subs):
            issues.append(f"{label}: some khata sub-totals could not be read and are not in its total")
        return {"mut": mut, "total": sub_sum, "kind": f"no grand total printed - sum of {len(sub)} khata sub-total(s)",
                "subs": subs, "unit": unit}
    issues.append(f"{label}: no total or khata sub-totals found")
    return {"mut": mut, "total": None, "kind": "", "subs": subs, "unit": ""}


def heal(srv, models, log):
    h = srv.heal_models(models)
    if h["action"] != "none":
        log(f"    model check: {h['action']} - {h['detail']}")
    return h


# ------------------------------------------------------- pages without a grid --

VISION_SCHEMA = {"type": "object", "properties": {"entries": {"type": "array", "items": {"type": "object", "properties": {
    "mutation_no": {"type": "string"}, "khata": {"type": "string"},
    "khasras": {"type": "array", "items": {"type": "string"}},
    "areas": {"type": "array", "items": {"type": "string"}}, "total": {"type": "string"}},
    "required": ["mutation_no", "khata", "khasras", "areas", "total"]}}}, "required": ["entries"]}
# No example values in the prompt: the model copies them into its answer when it can't read the page.
VISION_PROMPT = (
    "This is a page of a Pakistani land mutation register (Register Dakhil Kharij / Intiqal), a ruled form "
    "in Urdu. Columns run right to left: the rightmost holds the mutation number, the next the khewat/khata "
    "numbers; the leftmost two hold the khasra numbers (numbers joined by slashes) and the area of each "
    "(three numbers joined by dashes: Kanal, Marla, Sarsai), sometimes ending with a total area. Transcribe "
    "only what is written in those columns, for every entry on the page, as "
    '{"entries": [{"mutation_no": "", "khata": "", "khasras": [], "areas": [], "total": ""}]}. '
    "If you cannot read a value, leave it empty. Never guess, and never number the rows yourself.")


def vision_page(page, rot, ask_json):
    """A page whose ruled grid can't be found (handwritten, damaged, an unknown form): the vision model
    reads the whole upright page into the same records as mutation.read_page, flagged for checking."""
    z = 2000 / max(page.rect.width, page.rect.height)
    png = page.get_pixmap(matrix=pymupdf.Matrix(z, z).prerotate(rot), alpha=False).tobytes("png")
    try:
        import json
        entries = json.loads(ask_json(VISION_PROMPT, png, VISION_SCHEMA)).get("entries", [])[:12]
    except Exception as e:
        return {"ok": False, "records": [], "error": f"read without a grid failed ({type(e).__name__})"}
    # An answer numbering its entries 1, 2, 3 ... as mutation numbers is the model counting rows, not
    # reading them: the page couldn't be read.
    muts = [str(e.get("mutation_no") or "").strip() for e in entries]
    if len(muts) >= 3 and muts == [str(i) for i in range(1, len(muts) + 1)]:
        return {"ok": False, "records": [], "error": "the page could not be read without its grid (handwritten or unclear)"}
    records = []
    for e in entries:
        khasras = mutation.parse_khasra("\n".join(map(str, e.get("khasras") or []))).split(", ")
        khasras = [k for k in khasras if k]
        areas = mutation.KMS.findall(mutation._norm("\n".join(map(str, e.get("areas") or []))))
        items, col17_total, _ = mutation.pair_khasra_areas(khasras, areas)
        total = _kms(mutation._norm(e.get("total") or ""))
        col17_total = col17_total or ("-".join(map(str, (total[0], total[1], _fmt3(total[2])))) if total else "")
        mut = mutation.parse_award(e.get("mutation_no") or "")
        nums = mutation._plain_numbers(e.get("khata") or "")
        khata = re.sub(r"\s+", " ", str(e.get("khata") or "")).strip() if nums else ""
        if not (mut or khata or items or col17_total):
            continue
        records.append({"award": mut, "khata": khata, "katat": ", ".join(khasras), "kanal": "", "marla": "",
                        "sarsai": "", "continuation": False, "confidence": "low", "flags": ["vision"],
                        "notes": "read from the whole page without its grid - check", "khasra_only": False,
                        "total_only": not items and bool(col17_total) and not khata,
                        "khasra_items": items, "col17_total": col17_total, "cells": {}})
    return {"ok": True, "records": records, "error": "", "vision": True}


# ---------------------------------------------------------- which documents --

def classify(path, srv, log=lambda m: None):
    """What a PDF is, by its name - or, for names that don't say (phone scans like "CamScanner ..."), by
    its first pages: a register grid -> "mutation"; the word Award in its text -> "award"; else "other"."""
    name = os.path.basename(path)
    low = name.lower()
    if "correction" in low:
        return "other"
    if is_mutation(name):
        return "mutation"
    if award_batch.is_award(name):
        return "award"
    try:
        doc = pymupdf.open(path)
        for n in range(min(2, doc.page_count)):
            rot = srv.detect_orientation(doc, n)[0]
            if mutation.locate_grid(doc[n], rot):
                return "mutation"
        cmd = srv.tesseract_cmd()
        if cmd and doc.page_count:
            p = doc[0]
            png = p.get_pixmap(matrix=pymupdf.Matrix(200 / 72, 200 / 72), colorspace=pymupdf.csGRAY, alpha=False).tobytes("png")
            if re.search(r"\baward\b", " ".join(w["text"] for w in award.tesseract(cmd, png, psm=3)), re.I):
                return "award"
    except Exception as e:
        log(f"  could not check {name}: {e}")
    return "other"


# -------------------------------------------------------------------- word --

def _xml(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def docx_bytes(paragraphs):
    """A minimal Word document. paragraphs: [[(text, bold) or (text, bold, points), ...], ...] (an empty
    list is a blank line). Written directly (three XML parts in a zip) so the tool needs no extra package."""
    import io
    import zipfile

    def run(text, bold, size=None):
        props = ("<w:b/>" if bold else "") + (f'<w:sz w:val="{int(size * 2)}"/>' if size else "")
        props = f"<w:rPr>{props}</w:rPr>" if props else ""
        return f'<w:r>{props}<w:t xml:space="preserve">{_xml(text)}</w:t></w:r>'

    body = "".join(f"<w:p>{''.join(run(*r) for r in p)}</w:p>" for p in paragraphs)
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
                '</w:sectPr></w:body></w:document>')
    types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             '<Override PartName="/word/document.xml" '
             'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
    return buf.getvalue()


def area_words(t, unit):
    """(K, M, S) -> "149 K – 17 M" (with the third figure only when it isn't 0)."""
    if not t:
        return "not found"
    third = f" – {_fmt3(t[2])} {'Sq ft' if unit == 'Sq ft' else 'S'}" if t[2] else ""
    return f"{t[0]} K – {t[1]} M{third}"


def summary_paragraphs(rows, place, totals):
    """District / Tehsil / Village, then a section per mutation: each khata (in order of first appearance)
    with its khasras, and the mutation's total area."""
    out = [[("District: ", True), (place.get("district", ""), False)],
           [("Tehsil: ", True), (place.get("tehsil", ""), False)],
           [("Village/Mouza: ", True), (place.get("mouza", ""), False)]]
    for mt in totals:
        khatas = {}
        for r in rows:
            if r["mut"] == mt["mut"]:
                khatas.setdefault(r["khata"] or "(khata not read)", []).append(r["khasra"])
        out += [[], [(f"Mutation {mt['mut']}" if mt["mut"] else "Mutation", True, 14)], []]
        out += [[(f"{khata}: ", True), (", ".join(khasras), False)] for khata, khasras in khatas.items()]
        out += [[("Total Area: " + area_words(mt["total"], mt["unit"]), True)]]
    return out


# ------------------------------------------------------------------- excel --

def write_outputs(path, res, place, place_src):
    """Excel file next to the document; plus a review note when it is unreadable or too incomplete.
    -> (xlsx path or "", review note path or "", table rows as saved)"""
    folder = os.path.dirname(path)
    review = ""
    reasons = []
    if res["unreadable"]:
        reasons.append("no khasra rows could be read from the document")
    elif res["missing"] > MISSING_LIMIT:
        reasons.append(f"about {res['missing']:.0%} of the data is missing (limit {MISSING_LIMIT:.0%})")
    if res["failed_pages"]:
        reasons.append(f"{res['failed_pages']} of {res['pages']} page(s) could not be read")
    if reasons and (res["unreadable"] or res["missing"] > MISSING_LIMIT):
        review = review_name(path)
        with open(review, "w", encoding="utf-8") as f:
            f.write(f"REVIEW THIS DOCUMENT MANUALLY\n\nDocument: {path}\nChecked: "
                    f"{datetime.datetime.now():%Y-%m-%d %H:%M}\n\nWhy:\n" + "".join(f"- {r}\n" for r in reasons)
                    + ("\nNotes from the extraction:\n" + "".join(f"- {t}\n" for t in res["issues"]) if res["issues"] else ""))
    if res["unreadable"]:
        return "", review, []
    rows = [list(r) for r in res["table"]] + total_rows(res["totals"])
    info = [("Source file", path), ("District / Tehsil / Mouza from", place_src or "not found - fill them in"),
            ("Pages", res["pages"]), ("Khasras listed", len(res["rows"]))]
    for mt in res["totals"]:
        t = mt["total"]
        info.append((f"Total, mutation {mt['mut']}" if mt["mut"] else "Total",
                     f"{fmt_kms(t)} ({mt['kind']}; third figure read as {mt['unit'] or 'Sarsai'})" if t else "not found"))
    info += [("Data missing (estimate)", f"{res['missing']:.0%}"),
             ("Extracted", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))]
    counts = {}
    for t in res["issues"]:  # the same note for many entries is listed once, with a count
        counts[t] = counts.get(t, 0) + 1
    issues = [t if c == 1 else f"{t} (x{c})" for t, c in counts.items()]
    if review:
        issues.insert(0, "REVIEW MANUALLY: " + "; ".join(reasons))
    if not place_src:
        issues.append("No award document found in this folder (or above it) for District / Tehsil / Mouza")
    xlsx = os.path.join(folder, output_name(path, res["numbers"], place))
    with open(xlsx, "wb") as f:
        f.write(award.build_xlsx(rows, info, issues, headers=HEADERS, widths=WIDTHS, text_cols=(1, 5, 6),
                                 title="Mutation Land"))
    with open(re.sub(r"\.xlsx$", ".docx", xlsx), "wb") as f:  # the same land list as a Word summary
        f.write(docx_bytes(summary_paragraphs(res["rows"], place, res["totals"])))
    return xlsx, review, rows


def fmt_kms(t):
    return f"{t[0]}-{t[1]}-{_fmt3(t[2])}" if t else "not found"


def total_rows(totals):
    """Closing rows, one block per mutation: its khata sub-totals (only when no grand total was printed),
    then its Total."""
    out = []
    for mt in totals:
        t = mt["total"]
        if t and not mt["kind"].startswith("grand"):
            for mut, khata, area, page in mt["subs"]:
                a = _kms(area)
                if a:
                    out.append(["", mut, "", "", "", khata, "Sub-total", a[0], a[1], _fmt3(a[2]), page])
        out.append(["", mt["mut"], "", "", "", "", "Total", *(t[:2] if t else ("", "")), _fmt3(t[2]) if t else "", ""])
    return out


# --------------------------------------------------------------------- run --

def run(srv, root, models, ask, ask_json, log, redo=False, stop=lambda: False, step=None, live=None):
    """Convert every mutation document under `root`. ask(model, prompt, png) reads register cells;
    ask_json(prompt, png, schema) reads the award's place names; live(doc, rows, page, pages) streams the
    table being built. -> {"done", "review", "skipped", "failed", "stopped"}"""
    groups = find_documents(root)
    fh = open(os.path.join(root, "Mutation Batch Log.txt"), "a", encoding="utf-8")

    def say(msg):
        line = f"{datetime.datetime.now():%H:%M:%S}  {msg}"
        log(line)
        fh.write(line + "\n")
        fh.flush()

    unnamed = sum(len(u) for _f, _m, u in groups)
    say(f"=== Mutation batch over {root}: {sum(len(m) for _f, m, _u in groups)} mutation document(s) by name"
        + (f", {unnamed} other PDF(s) to check" if unnamed else "") + f" in {len(groups)} folder(s)")
    n = {"done": 0, "review": 0, "skipped": 0, "failed": 0, "stopped": False}
    places, found_awards, docs = {}, [], []
    try:
        if groups:
            h = heal(srv, list(models), say)  # start from healthy models
            if not h["ok"]:
                say(f"!! The models are not reading correctly ({h['detail']}) - results may be incomplete")
        for folder, files, others in groups:
            docs += [(folder, f) for f in files]
            for f in others:  # phone scans etc.: what they are is told from their pages
                if stop():
                    break
                if not redo and is_done(f, os.listdir(folder)):
                    docs.append((folder, f))  # converted before: skipped below without checking it again
                    continue
                kind = classify(f, srv, say)
                say(f"{award_batch.folder_label(folder, root)}\\{os.path.basename(f)}: looks like "
                    + {"mutation": "a mutation register", "award": "an award (used for District / Tehsil / Mouza)",
                       "other": "neither - left alone"}[kind])
                if kind == "mutation":
                    docs.append((folder, f))
                elif kind == "award":
                    found_awards.append(f)
        for i, (folder, path) in enumerate(docs):
            if stop():
                n["stopped"] = True
                say("Stopped by the user.")
                break
            name = os.path.basename(path)
            rel = os.path.join(award_batch.folder_label(folder, root), name)
            if step:
                step(i, len(docs), rel)
            if not redo and is_done(path, os.listdir(folder)):
                say(f"{rel}: already converted (or marked for review), skipped")
                n["skipped"] += 1
                continue
            say(rel)
            t0 = time.time()
            try:
                if folder not in places:
                    places[folder] = place_from_award(folder, root, ask_json, say, found_awards)
                place, src = places[folder]
                res = read_document(path, srv, ask, models, place, say, stop=stop, ask_json=ask_json,
                                    live=(lambda rows, p, pages: live(name, rows, p, pages)) if live else None)
                xlsx, review, saved = write_outputs(path, res, place, src)
                if live and saved:
                    live(name, saved, res["pages"], res["pages"])  # as saved, with the total at the end
                tots = ", ".join((f"{mt['mut']}: " if mt["mut"] else "") + fmt_kms(mt["total"]) for mt in res["totals"])
                say(f"  -> {os.path.basename(xlsx) if xlsx else 'no Excel file'}: {len(res['rows'])} khasras, "
                    f"total{'s' if len(res['totals']) > 1 else ''} {tots}, {res['missing']:.0%} missing"
                    f"{' - REVIEW MANUALLY note written' if review else ''}, {time.time() - t0:.0f}s")
                n["review" if review else "done"] += 1
            except Exception as e:
                n["failed"] += 1
                say(f"  !! failed: {e}")
                fh.write(traceback.format_exc() + "\n")
        say(f"=== Finished: {n['done']} converted, {n['review']} need manual review, {n['skipped']} skipped, "
            f"{n['failed']} failed")
    finally:
        fh.close()
    return n
