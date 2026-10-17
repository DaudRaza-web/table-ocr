# Table OCR

A fully offline tool that turns tables in scanned documents, PDFs and images into Excel files.
Open a document, drag a box around a table, and a local vision model reads it into an editable grid.
Fix any cells, add a title, then download as `.xlsx`. Nothing leaves your machine.

## Features

- Open PDFs or images (PNG, JPG, TIFF, …) by picking a file or dragging it in
- Page-by-page viewer with zoom
- Drag to select the table region on the page
- Extraction by a local vision model running in [Ollama](https://ollama.com) (default: `qwen2.5vl:7b`)
- Editable grid: edit cells, add or remove rows and columns, mark the first row as a header
- **Append rows** for tables that continue across pages (a repeated header row is skipped)
- Optional table title, placed above the table in the spreadsheet
- `.xlsx` export with real numbers (`12,450`, `12.3%`, `$1,200.50`, `(9,870)` keep their formatting), or copy as tab-separated text to paste into Excel
- Optional LAN mode so other devices on your network can use it

## Requirements

- Python 3.10+
- [Ollama](https://ollama.com) with a vision model. On Windows it can run natively or inside WSL;
  the server talks to `http://localhost:11434`, which WSL2 forwards automatically.
- A GPU is recommended. `qwen2.5vl:7b` needs ~6 GB of VRAM; the first extraction loads the model (~30–60 s), later ones take a few seconds.

Pull the model once (the only step that needs internet besides installing the Python packages):

```bash
ollama pull qwen2.5vl:7b
```

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
| `TABLE_OCR_MODEL` | `qwen2.5vl:7b` | Model selected by default |
| `OLLAMA_URL` | `http://localhost:11434` | Ollama endpoint |
| `TABLE_OCR_WSL_DISTRO` | default distro | WSL distro to start Ollama in |

Pass `--no-browser` to skip opening the browser on start.

## How it works

`server.py` is a small standard-library HTTP server. It renders pages with PyMuPDF, crops the selected region at high resolution,
and sends it to Ollama's `/api/chat` with a JSON schema, so the model returns a clean list of rows. The grid lives in
`static/index.html` (plain HTML/JS, no external libraries), and openpyxl writes the `.xlsx`.

`sample_report.pdf` is a two-page example with a table that continues onto the second page. Use it to try the tool.
