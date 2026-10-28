"""Batch award extraction: every folder under a root, one award per folder -> an Excel land table.

    python award_batch.py ["D:\\...\\M3 Data"] [--redo] [--pages N] [--file award.pdf] [--out folder]

Also run from the web UI's Award tab ("Convert folder"), through run().

Each folder's award is read with the table extractor (qwen2.5vl only, every other model unloaded so it
has the machine's memory to itself), the land table is built (square carried down and joined to the
khasra, total K-M from the document) and saved next to the award as "Award <no> <mouza> - Land.xlsx".
When a folder holds more than one award, the user is asked which to use; after 5 minutes without an
answer the latest award (highest year, then award number) is taken. Correction slips are not awards.
Folders that already have the Excel file are skipped unless --redo is given.
"""
import argparse
import base64
import datetime
import os
import re
import sys
import time
import traceback

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
import award  # noqa: E402

DEFAULT_ROOT = r"D:\Danish Enterprises\Motorway\M3 Data"
PICK_TIMEOUT = 5 * 60
OPTIONS = {"temperature": 0, "num_ctx": 16384, "num_predict": 8192}


# ------------------------------------------------------------------- model --
# `srv` is the server module (Ollama access, prompts, Tesseract); passed in so the web server can run
# batches without importing itself a second time.

def one_model(srv, model):
    """Unload every model except the table reader, then load it and keep it loaded."""
    srv.ensure_ollama()
    try:
        loaded = [m["name"] for m in srv.ollama_request("/api/ps").get("models", [])]
    except Exception:
        loaded = []
    srv.unload_models([m for m in loaded if m != model])
    srv.ollama_request("/api/generate", {"model": model, "keep_alive": "60m"}, timeout=600)


def askers(srv, model):
    """-> (ask_table(png) -> rows, chat(prompt, png, schema) -> str)"""
    def chat(prompt, png, schema):
        payload = {"model": model, "stream": False, "format": schema, "keep_alive": "60m", "options": OPTIONS,
                   "messages": [{"role": "user", "content": prompt, "images": [base64.b64encode(png).decode()]}]}
        for attempt in range(3):
            try:
                return srv.ollama_request("/api/chat", payload, timeout=900).get("message", {}).get("content", "")
            except Exception:
                if attempt == 2:
                    raise
                srv.ensure_ollama()
                time.sleep(5)

    return (lambda png: srv.parse_rows(chat(srv.PROMPT, png, srv.SCHEMA))), chat


# ------------------------------------------------------------------ awards --

def is_award(name):
    n = name.lower()
    return n.endswith(".pdf") and "award" in n and "correction" not in n


def is_output(name):
    return name.startswith("Award") and name.endswith("- Land.xlsx")


def award_key(path):
    """Latest first: year, then award number, then file date."""
    m = award.AWARD_NO.search(os.path.basename(path))
    year = int(m.group(2)) if m and m.group(2) else 0
    no = int(m.group(1)) if m else 0
    return (year, no, os.path.getmtime(path))


def find_folders(root):
    """-> [(folder, [award pdfs], already_done)] for every folder under root that holds an award."""
    out = []
    for folder, dirs, names in os.walk(root):
        dirs.sort()
        files = [os.path.join(folder, n) for n in sorted(names) if is_award(n)]
        if files:
            out.append((folder, files, any(is_output(n) for n in names)))
    return out


def pick_window(folder, files):
    """Desktop window asking which award to use; the latest is taken after PICK_TIMEOUT. -> (path, how)"""
    latest = max(files, key=award_key)
    try:
        import tkinter as tk
    except ImportError:
        return latest, "auto (no window available)"
    choice = {"path": latest, "how": f"auto - no answer in {PICK_TIMEOUT // 60} minutes, latest award taken"}
    win = tk.Tk()
    win.title("Which award?")
    win.attributes("-topmost", True)
    tk.Label(win, text=f"This folder has {len(files)} awards:\n{folder}\n\nWhich one should be used?",
             justify="left", padx=14, pady=10).pack(anchor="w")
    var = tk.StringVar(value=latest)
    for f in sorted(files, key=award_key, reverse=True):
        tk.Radiobutton(win, text=os.path.basename(f) + ("   (latest)" if f == latest else ""), variable=var,
                       value=f, anchor="w", padx=14).pack(fill="x")
    timer = tk.Label(win, padx=14, pady=6, fg="#555")
    timer.pack(anchor="w")
    end = time.time() + PICK_TIMEOUT

    def ok():
        choice.update(path=var.get(), how="chosen by user")
        win.destroy()

    def tick():
        left = int(end - time.time())
        if left <= 0:
            win.destroy()
            return
        timer.config(text=f"The latest award is used automatically in {left // 60}:{left % 60:02d}")
        win.after(1000, tick)

    tk.Button(win, text="Use this award", command=ok, padx=10).pack(pady=(0, 12))
    win.protocol("WM_DELETE_WINDOW", win.destroy)  # closing the window = take the latest
    tick()
    win.mainloop()
    return choice["path"], choice["how"]


# ------------------------------------------------------------------- excel --

def write_xlsx(path, src, res, how):
    total = res["total"]
    info = [("Source file", src), ("Award chosen", how),
            ("Pages with the land table", ", ".join(map(str, res["pages"])) or "none"),
            ("Khasras listed", len(res["rows"])),
            ("Total (from the document)", f"{total[0]} K - {total[1]} M" if total else "not found"),
            ("Extracted", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))]
    rows = award.table_rows(res["award"], res["place"], res["rows"], total)
    with open(path, "wb") as f:
        f.write(award.build_xlsx(rows, info, res["issues"]))


def out_name(res):
    no = (res["award"] or "").replace("/", "-")
    mouza = re.sub(r'[<>:"/\\|?*]', "-", res["place"].get("mouza", "")).strip()
    return " ".join(x for x in ("Award", no, mouza, "- Land.xlsx") if x)


# --------------------------------------------------------------------- run --

def process(srv, model, src, how, out_dir, max_pages, log):
    ask_table, chat = askers(srv, model)
    t0 = time.time()
    res = award.extract(src, ask_table, chat, srv.tesseract_cmd(), max_pages=max_pages,
                        progress=lambda n, tot: log(f"    page {n + 1}/{tot}"))
    path = os.path.join(out_dir, out_name(res))
    write_xlsx(path, src, res, how)
    tot = f"{res['total'][0]}-{res['total'][1]}" if res["total"] else "no total"
    log(f"  -> {os.path.basename(path)}: {len(res['rows'])} khasras, total {tot}, "
        f"{len(res['issues'])} issue(s), {time.time() - t0:.0f}s")
    return path


def run(srv, root, model, pick, log, redo=False, stop=lambda: False, max_pages=None, out=None, step=None):
    """Convert the award in every folder under `root`.
    pick(folder, files) -> (path, how) when a folder has several awards; log(text); stop() -> True to halt
    after the current award; step(i, n, folder) reports progress. -> {"done", "skipped", "failed", "stopped"}"""
    folders = find_folders(root)
    fh = open(os.path.join(root, "Award Batch Log.txt"), "a", encoding="utf-8")

    def say(msg):
        line = f"{datetime.datetime.now():%H:%M:%S}  {msg}"
        log(line)
        fh.write(line + "\n")
        fh.flush()

    say(f"=== Award batch over {root}: {len(folders)} folder(s) with awards")
    n = {"done": 0, "skipped": 0, "failed": 0, "stopped": False}
    try:
        if any(redo or not d for _f, _a, d in folders):
            say("Loading the table model (other models unloaded) ...")
            one_model(srv, model)
        for i, (folder, files, done) in enumerate(folders):
            if stop():
                n["stopped"] = True
                say("Stopped by the user.")
                break
            if step:
                step(i, len(folders), folder)
            rel = os.path.relpath(folder, root)
            if done and not redo:
                say(f"{rel}: already has its Excel file, skipped")
                n["skipped"] += 1
                continue
            if len(files) > 1:
                say(f"{rel}: {len(files)} awards - asking which to use")
                src, how = pick(folder, files)
            else:
                src, how = files[0], "only award in the folder"
            say(f"{rel}: {os.path.basename(src)} ({how})")
            try:
                process(srv, model, src, how, out or folder, max_pages, say)
                n["done"] += 1
            except Exception as e:
                n["failed"] += 1
                say(f"  !! failed: {e}")
                fh.write(traceback.format_exc() + "\n")
        say(f"=== Finished: {n['done']} converted, {n['skipped']} skipped, {n['failed']} failed")
    finally:
        fh.close()
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root", nargs="?", default=DEFAULT_ROOT)
    ap.add_argument("--redo", action="store_true", help="re-extract folders that already have the Excel file")
    ap.add_argument("--pages", type=int, help="read only the first N pages (testing)")
    ap.add_argument("--file", help="one award PDF instead of the whole tree (testing)")
    ap.add_argument("--out", help="save the Excel file here instead of next to the award")
    a = ap.parse_args()
    sys.argv = sys.argv[:1]  # server.py reads sys.argv at import
    import server
    model = server.DEFAULT_MODEL
    if a.file:
        print("Loading the table model (other models unloaded) ...", flush=True)
        one_model(server, model)
        process(server, model, a.file, "single file", a.out or os.path.dirname(a.file), a.pages,
                lambda m: print(m, flush=True))
        return
    run(server, a.root, model, pick_window, lambda m: print(m, flush=True), redo=a.redo, max_pages=a.pages, out=a.out)


if __name__ == "__main__":
    main()
