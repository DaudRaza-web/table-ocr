"""Batch mutation extraction: every mutation PDF under a root -> an Excel table next to it.

Run from the web UI's "Mutation batch" tab (server.py), through run().

Each mutation document (Register Dakhil Kharij pages) is read with the register reader (mutation.py,
khasra mode). One row per khasra: Mutation No, Khata Khatuni, Khasra No.; the areas are not listed
per khasra, only the document's total at the end: its grand total (میزان کل) when printed, otherwise
the khatas' printed sub-totals added up. District / Tehsil / Mouza come from the award document in the
same folder (its converted Excel file if there is one), not from the mutation.

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


def output_name(path, numbers, place):
    nos = " & ".join(numbers) or re.sub(r"\.pdf$", "", os.path.basename(path), flags=re.I)
    mouza = re.sub(r'[<>:"/\\|?*]', "-", place.get("mouza", "")).strip()
    return " ".join(x for x in ("Mutation", nos, mouza, "- Land.xlsx") if x)


def review_name(path):
    return re.sub(r"\.pdf$", "", path, flags=re.I) + " - REVIEW MANUALLY.txt"


def numbers_in_name(path):
    """"Mutation Nos. 3272 & 3273 Village ..." -> ["3272", "3273"]"""
    m = re.search(r"mutation\s*nos?\.?\s*([\d\s&,and]+)", os.path.basename(path), re.I)
    return re.findall(r"\d+", m.group(1)) if m else []


def find_documents(root):
    """-> [(folder, [mutation pdfs])] for every folder under root that holds mutation documents."""
    out = []
    for folder, dirs, names in os.walk(root):
        dirs.sort()
        files = [os.path.join(folder, n) for n in sorted(names) if is_mutation(n)]
        if files:
            out.append((folder, files))
    return out


def is_done(path, names):
    """A document is done when its Excel file or its review note exists."""
    stem = re.sub(r"\.pdf$", "", os.path.basename(path), flags=re.I)
    return os.path.basename(review_name(path)) in names or any(
        n.startswith("Mutation") and n.endswith("- Land.xlsx") and _same_doc(n, stem) for n in names)


def _same_doc(xlsx, stem):
    nums = numbers_in_name(stem + ".pdf")
    return bool(nums) and all(re.search(rf"\b{n}\b", xlsx) for n in nums)


# ------------------------------------------------------------------- place --

def place_from_award(folder, root, ask_json, log):
    """District / Tehsil / Mouza from an award in this folder (or the folders above it, up to root):
    its converted Excel file if there is one, else read from the award PDF. -> (place, source)"""
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
        awards = [os.path.join(here, n) for n in names if award_batch.is_award(n)]
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


def read_document(path, srv, ask, models, place, log, live=None, stop=lambda: False):
    """One mutation PDF -> {"rows", "total", "total_kind", "subtotals", "issues", "missing", "unreadable",
    "numbers", "pages"}. live(rows, page, pages) is called after each page with the table so far."""
    doc = pymupdf.open(path)
    named = numbers_in_name(path)
    rows, subtotals, issues, totals_only = [], [], [], []
    expected = missing = failed = 0
    last_mut, last_khata = (named[0] if len(named) == 1 else ""), {}
    since_check = 0
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
                                         plot_field="khasra", keep_totals=True)
            except Exception as e:
                res = {"ok": False, "records": [], "error": f"{type(e).__name__}: {e}", "crashed": True}
            bad = res.get("crashed") or (res["ok"] and any(
                f in ("award", "katat") for rec in res["records"] for f in rec.get("flags", [])))
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
            continue
        for rec in res["records"]:
            mut = rec["award"] or last_mut
            if rec["award"]:
                last_mut = rec["award"]
            if rec.get("total_only") or (not rec["khasra_items"] and rec["col17_total"] and not rec["khasra_only"]):
                totals_only.append((rec["col17_total"], n + 1))  # a total with no khasras: maybe the grand total
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
            seen = {r["khasra"] for r in rows if r["mut"] == mut} if rec["khasra_only"] else set()
            for it in rec["khasra_items"]:
                if it["khasra"] in seen:
                    continue  # a continuation page repeating khasras already listed
                rows.append({"mut": mut, "khata": khata, "khasra": it["khasra"], "page": n + 1})
                expected += 3
                missing += (not mut) + (not khata)
            if rec["col17_total"]:
                subtotals.append((mut, khata, rec["col17_total"], n + 1))
            for note in filter(None, (rec.get("notes") or "").split("; ")):
                if not note.startswith("column 14"):
                    issues.append(f"Page {n + 1}: {note}")
        if live:
            live(table(), n + 1, doc.page_count)

    # Total: the grand total printed at the end (a column 17 area with no khasras after the last entry,
    # at least the khatas' sub-totals added up), else the printed sub-totals added up.
    total, kind = None, ""
    sub = [a for a in (_kms(s[2]) for s in subtotals) if a]
    sub_sum, unit = add_areas(sub) if sub else (None, "")
    if totals_only:
        g = _kms(totals_only[-1][0])
        per = 272 if unit == "Sq ft" or (g and g[2] >= 9) else 9
        size = lambda a: a[0] * 20 * per + a[1] * per + a[2]  # noqa: E731
        if g and (not sub or size(g) + 0.01 >= size(sub_sum)):
            unit = unit or ("Sq ft" if per == 272 else "Sarsai")
            total, kind = g, f"grand total printed on page {totals_only[-1][1]}"
        elif g:
            issues.append(f"Page {totals_only[-1][1]}: area {totals_only[-1][0]} with no khasras is smaller than the "
                          "khata sub-totals added up, so it was not taken as the grand total")
    if total is None and sub:
        total, kind = sub_sum, f"no grand total printed - sum of {len(sub)} khata sub-total(s)"
        if len(sub) < len(subtotals):
            issues.append("Some khata sub-totals could not be read and are not in the total")
    if total is None:
        issues.append("No total or khata sub-totals found in the document")
    numbers = list(dict.fromkeys(r["mut"] for r in rows if r["mut"])) or named
    ratio = (missing / expected) if expected else 1.0
    page_ratio = failed / doc.page_count if doc.page_count else 1.0
    unreadable = not rows
    return {"rows": rows, "table": table(), "total": total, "total_kind": kind, "unit": unit,
            "subtotals": subtotals, "issues": issues, "missing": max(ratio, page_ratio), "unreadable": unreadable,
            "failed_pages": failed, "numbers": numbers, "pages": doc.page_count}


def heal(srv, models, log):
    h = srv.heal_models(models)
    if h["action"] != "none":
        log(f"    model check: {h['action']} - {h['detail']}")
    return h


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
    rows = [list(r) for r in res["table"]]
    t = res["total"]
    if t and not res["total_kind"].startswith("grand"):
        for mut, khata, area, page in res["subtotals"]:
            a = _kms(area)
            if a:
                rows.append(["", mut, "", "", "", khata, "Sub-total", a[0], a[1], _fmt3(a[2]), page])
    rows.append(["", "", "", "", "", "", "Total", *(t[:2] if t else ("", "")), _fmt3(t[2]) if t else "", ""])
    info = [("Source file", path), ("District / Tehsil / Mouza from", place_src or "not found - fill them in"),
            ("Pages", res["pages"]), ("Khasras listed", len(res["rows"])),
            ("Total", f"{t[0]}-{t[1]}-{_fmt3(t[2])} ({res['total_kind']}; third figure read as {res['unit'] or 'Sarsai'})"
             if t else "not found"),
            ("Data missing (estimate)", f"{res['missing']:.0%}"),
            ("Extracted", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))]
    issues = list(res["issues"])
    if review:
        issues.insert(0, "REVIEW MANUALLY: " + "; ".join(reasons))
    if not place_src:
        issues.append("No award document found in this folder (or above it) for District / Tehsil / Mouza")
    xlsx = os.path.join(folder, output_name(path, res["numbers"], place))
    with open(xlsx, "wb") as f:
        f.write(award.build_xlsx(rows, info, issues, headers=HEADERS, widths=WIDTHS, text_cols=(1, 5, 6),
                                 title="Mutation Land"))
    return xlsx, review, rows


# --------------------------------------------------------------------- run --

def run(srv, root, models, ask, ask_json, log, redo=False, stop=lambda: False, step=None, live=None):
    """Convert every mutation document under `root`. ask(model, prompt, png) reads register cells;
    ask_json(prompt, png, schema) reads the award's place names; live(doc, rows, page, pages) streams the
    table being built. -> {"done", "review", "skipped", "failed", "stopped"}"""
    groups = find_documents(root)
    docs = [(folder, f) for folder, files in groups for f in files]
    fh = open(os.path.join(root, "Mutation Batch Log.txt"), "a", encoding="utf-8")

    def say(msg):
        line = f"{datetime.datetime.now():%H:%M:%S}  {msg}"
        log(line)
        fh.write(line + "\n")
        fh.flush()

    say(f"=== Mutation batch over {root}: {len(docs)} mutation document(s) in {len(groups)} folder(s)")
    n = {"done": 0, "review": 0, "skipped": 0, "failed": 0, "stopped": False}
    places = {}
    try:
        if docs:
            h = heal(srv, list(models), say)  # start from healthy models
            if not h["ok"]:
                say(f"!! The models are not reading correctly ({h['detail']}) - results may be incomplete")
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
                    places[folder] = place_from_award(folder, root, ask_json, say)
                place, src = places[folder]
                res = read_document(path, srv, ask, models, place, say, stop=stop,
                                    live=(lambda rows, p, pages: live(name, rows, p, pages)) if live else None)
                xlsx, review, saved = write_outputs(path, res, place, src)
                if live and saved:
                    live(name, saved, res["pages"], res["pages"])  # as saved, with the total at the end
                t = res["total"]
                say(f"  -> {os.path.basename(xlsx) if xlsx else 'no Excel file'}: {len(res['rows'])} khasras, total "
                    f"{f'{t[0]}-{t[1]}-{_fmt3(t[2])}' if t else 'not found'}, {res['missing']:.0%} missing"
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
