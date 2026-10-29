"""Rewrite a mutation document's Excel and Word files from corrected values (checked by eye).

    python fix_doc.py <corrections.json>

corrections.json:
{"pdf": "D:\\...\\CamScanner 29-10-2026 15.49.pdf",
 "place": {"district": "", "tehsil": "", "mouza": ""},          (optional: else from the folder's award)
 "rows": [{"mut": "525", "khata": "103, 99, 141", "khasra": "130/1", "page": 1}, ...],
 "totals": [{"mut": "525", "total": "57-0-0", "kind": "printed under the table, page 2"}, ...],
 "notes": ["anything still uncertain"]}

The tool's earlier files for that PDF (its Excel, Word and any REVIEW MANUALLY note) are replaced.
"""
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(HERE)))
import mutation_batch as mb  # noqa: E402


def main():
    with open(sys.argv[1], encoding="utf-8") as f:
        c = json.load(f)
    pdf = c["pdf"]
    folder = os.path.dirname(pdf)
    stem = mb._stem(pdf)
    nums = mb.numbers_in_name(pdf)
    place, src = c.get("place"), "given in the corrections"
    for n in sorted(os.listdir(folder)):  # the tool's earlier outputs for this PDF
        mine = (f"({stem})" in n) or (nums and n.startswith("Mutation") and mb._same_doc(n, stem))
        if mine and n.endswith("- Land.xlsx") and not place:  # keep the place names the tool found
            from openpyxl import load_workbook
            wb = load_workbook(os.path.join(folder, n), read_only=True)
            for row in wb.worksheets[0].iter_rows(min_row=2, max_row=2, values_only=True):
                place = {"district": row[2] or "", "tehsil": row[3] or "", "mouza": row[4] or ""}
            src = [r[1] for r in wb["Review"].iter_rows(values_only=True) if r[0] == "District / Tehsil / Mouza from"]
            src = src[0] if src else "the tool's earlier file"
            wb.close()
        if (mine and n.endswith(("- Land.xlsx", "- Land.docx"))) or n == os.path.basename(mb.review_name(pdf)):
            os.remove(os.path.join(folder, n))
    if not place:
        place, src = mb.place_from_award(folder, folder, None, print)
    rows = [{"mut": str(r["mut"]), "khata": str(r["khata"]), "khasra": str(r["khasra"]), "page": r.get("page", "")}
            for r in c["rows"]]
    totals = []
    for t in c["totals"]:
        a = mb._kms(str(t.get("total") or ""))
        totals.append({"mut": str(t["mut"]), "total": a, "kind": t.get("kind") or "checked by eye", "subs": [],
                       "unit": "Sq ft" if a and a[2] >= 9 else "Sarsai"})
    table = [[i, r["mut"], place.get("district", ""), place.get("tehsil", ""), place.get("mouza", ""), r["khata"],
              r["khasra"], "", "", "", r["page"]] for i, r in enumerate(rows, 1)]
    res = {"rows": rows, "table": table, "totals": totals, "missing": 0.0, "unreadable": False, "failed_pages": 0,
           "pages": c.get("pages", ""), "numbers": list(dict.fromkeys(r["mut"] for r in rows if r["mut"])),
           "issues": ["Checked against the scan and corrected by eye (not by the tool)"] + list(c.get("notes", []))}
    xlsx, _review, _rows = mb.write_outputs(pdf, res, place, src)
    print("written", xlsx, re.sub(r"\.xlsx$", ".docx", xlsx))


if __name__ == "__main__":
    main()
