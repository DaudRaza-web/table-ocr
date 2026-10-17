# Table OCR

A fully offline tool that turns tables in scanned documents, PDFs and images into Excel files.
Open a document, drag a box around a table, and a local vision model reads it into an editable grid.
Fix any cells, add a title, then download as `.xlsx`. Nothing leaves your machine.

## Features

- Open PDFs or images (PNG, JPG, TIFF, …) by picking a file or dragging it in
- Page-by-page viewer with zoom
- **Auto-rotate**: sideways and upside-down pages are turned upright automatically (from the text layer in
  digital PDFs, or Tesseract's orientation detection for scans), plus ⟲ ⟳ buttons to rotate any page by hand
- Drag to select the table region on the page
- Extraction by a local vision model running in [Ollama](https://ollama.com) (default: `qwen2.5vl:7b`)
- Editable grid: edit cells, add or remove rows and columns, mark the first row as a header
- **Append rows** for tables that continue across pages (a repeated header row is skipped)
- Optional table title, placed above the table in the spreadsheet
- `.xlsx` export with real numbers (`12,450`, `12.3%`, `$1,200.50`, `(9,870)` keep their formatting), or copy as tab-separated text to paste into Excel
- Optional LAN mode so other devices on your network can use it
- **Translate tab**: Urdu, Punjabi (Shahmukhi and Gurmukhi) and Sindhi to English, using
  [TranslateGemma](https://ollama.com/library/translategemma) 12B. Paste text, or select a region of the page:
  it's transcribed into an editable "Original text" box first, so you can fix reading errors, then translated.
  The source language is auto-detected (Sindhi-only letters, Gurmukhi script, common Punjabi words).

## Requirements

- Python 3.10+
- [Ollama](https://ollama.com) with a vision model. On Windows it can run natively or inside WSL;
  the server talks to `http://localhost:11434`, which WSL2 forwards automatically.
- A GPU is recommended. `qwen2.5vl:7b` needs ~6 GB of VRAM; the first extraction loads the model (~30–60 s), later ones take a few seconds.

Pull the model once (the only step that needs internet besides installing the Python packages):

```bash
ollama pull qwen2.5vl:7b
ollama pull translategemma:12b   # for the Translate tab (~8 GB, needs ~10 GB VRAM)
```

For auto-rotating **scanned** pages, install Tesseract (optional; digital PDFs don't need it). The server finds it
on `PATH`, or on Windows inside WSL:

```bash
sudo apt install tesseract-ocr
```

Without Tesseract, scanned pages lying on their side are still turned upright, but which way round is a guess,
and upside-down scans aren't detected. (Vision LLMs turned out to be unreliable at telling upright from upside-down.)

If a pull stalls near the end, restart Ollama (`sudo systemctl restart ollama`) and pull again — it resumes.

Other vision models in Ollama (e.g. `granite3.2-vision`, `qwen3-vl`, `llama3.2-vision`) show up in the model dropdown.

## Run

**Windows:** double-click `start.bat`. The first run creates a `.venv` and installs `pymupdf` and `openpyxl`.
The tool then opens at <http://127.0.0.1:8765>.

**Other platforms / manual:**

```bash
python -m venv .venv
.venv/bin/pip install pymupdf openpyxl
.venv/bin/python server.py
```

On Windows the server also wakes WSL and starts `ollama serve` inside it if Ollama isn't reachable.

### Use it from other devices on your network

Run `start-lan.bat` (or `python server.py --lan`). The console prints addresses like `http://192.168.x.x:8765`.
You may need to allow the port through your firewall, e.g. on Windows (admin PowerShell):

```powershell
New-NetFirewallRule -DisplayName "Table OCR (LAN)" -Direction Inbound -Protocol TCP -LocalPort 8765 -RemoteAddress LocalSubnet -Action Allow -Profile Any
```

LAN mode has no authentication: anyone on your network can use the tool while it runs.

## Configuration

| Environment variable | Default | Purpose |
| --- | --- | --- |
| `TABLE_OCR_PORT` | `8765` | Port the web UI listens on |
| `TABLE_OCR_HOST` | `127.0.0.1` (`0.0.0.0` with `--lan`) | Interface to listen on |
| `TABLE_OCR_MODEL` | `qwen2.5vl:7b` | Table extraction model selected by default |
| `TABLE_OCR_TRANSLATE_MODEL` | `translategemma:12b` | Translation model selected by default |
| `OLLAMA_URL` | `http://localhost:11434` | Ollama endpoint |
| `TABLE_OCR_WSL_DISTRO` | default distro | WSL distro to start Ollama in |

Pass `--no-browser` to skip opening the browser on start.

## How it works

`server.py` is a small standard-library HTTP server. It renders pages with PyMuPDF, crops the selected region at high resolution,
and sends it to Ollama's `/api/chat` with a JSON schema, so the model returns a clean list of rows. The grid lives in
`static/index.html` (plain HTML/JS, no external libraries), and openpyxl writes the `.xlsx`.

`sample_report.pdf` is a two-page example with a table that continues onto the second page. Use it to try the tool.
