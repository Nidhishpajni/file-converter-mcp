# FileForge — File Converter MCP Server

<p align="center">
  <img src="https://img.shields.io/badge/MCP-Compatible-blue?style=for-the-badge&logo=anthropic" alt="MCP Compatible">
  <img src="https://img.shields.io/badge/Tools-83-brightgreen?style=for-the-badge" alt="83 Tools">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge" alt="MIT License">
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python" alt="Python 3.9+">
  <img src="https://img.shields.io/badge/Web_UI-Included-purple?style=for-the-badge" alt="Web UI">
</p>

<p align="center">
  <strong>The open-source ilovepdf.com for AI agents and humans alike.</strong><br>
  83 file conversion tools across 9 categories — usable as an MCP server by Claude, Cursor, and any MCP-compatible agent, <em>or</em> as a standalone web app.
</p>

---

## What is this?

Most AI agents can write and read files, but they can't *convert* them. **FileForge** fills that gap.

Add it to Claude, Cursor, Windsurf, or any MCP client and your AI can:

- *"Compress this PDF and protect it with a password"*
- *"Convert all images in this folder to WebP at 80% quality"*
- *"Extract tables from this PDF to Excel"*
- *"Turn this CSV into a Parquet file"*
- *"Merge these three PDFs and add page numbers"*
- *"OCR this scanned document and make it searchable"*

No API keys. No cloud. Everything runs locally.

---

## Features at a Glance

| Category | Tools | Highlights |
|----------|-------|-----------|
| **PDF Core** | 11 | Merge, split, compress, rotate, protect/unlock, watermark, page numbers, organize, repair |
| **PDF → Other** | 7 | PDF → Word, Excel, PowerPoint, Images, HTML, Text, Markdown |
| **Other → PDF** | 5 | Images, Word, Excel, PowerPoint, HTML → PDF |
| **OCR** | 6 | Scanned PDF → searchable PDF, image → text/Word/HTML/Markdown, 80+ languages |
| **Images** | 7 | PNG/JPG/WebP/BMP/GIF/TIFF conversion, resize, GIF frames, Base64, HEIC |
| **Documents** | 20 | DOCX/HTML/Markdown/TXT/PPTX/XLSX → any other doc format |
| **Data** | 18 | JSON/YAML/CSV/XML/TOML/XLSX/Parquet/SQLite/JSONL conversions |
| **Archives** | 6 | Create & extract ZIP, TAR.GZ, 7Z |
| **E-books** | 7 | EPUB → TXT/HTML/Markdown/DOCX/PDF, RTF → TXT/HTML |

**Total: 87 tools** (83 conversion tools + `get_server_info`)

---

## Installation

```bash
git clone https://github.com/Nidhishpajni/file-converter-mcp.git
cd file-converter-mcp
pip install -r requirements.txt
```

### Optional extras

```bash
# For HTML → PDF with full CSS support
pip install weasyprint

# For Office → PDF (also requires Microsoft Office on Windows)
pip install docx2pdf

# For HEIC/HEIF image support
pip install pillow-heif

# For e-book tools
pip install ebooklib striprtf

# For Parquet support
pip install pyarrow

# Pre-download OCR models (~200 MB)
python -c "import easyocr; easyocr.Reader(['en'])"
```

---

## Quick Start

### Add to Claude Code (CLI)

```bash
claude mcp add file-converter python /path/to/file-converter-mcp/file_converter_mcp_server.py
```

### Add to Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or
`%APPDATA%\Claude\claude_desktop_config.json` (Windows):

```json
{
  "mcpServers": {
    "file-converter": {
      "command": "python",
      "args": ["/absolute/path/to/file-converter-mcp/file_converter_mcp_server.py"]
    }
  }
}
```

### Add to Cursor / Windsurf / any MCP client

```json
{
  "mcpServers": {
    "file-converter": {
      "command": "python",
      "args": ["/absolute/path/to/file-converter-mcp/file_converter_mcp_server.py"]
    }
  }
}
```

### Run the Web UI

```bash
pip install fastapi uvicorn python-multipart
python web_app.py
# Open http://localhost:8080
```

### Start the MCP server manually

```bash
python file_converter_mcp_server.py              # stdio (default)
python file_converter_mcp_server.py --transport sse --port 8000
```

### Test with MCP Inspector

```bash
mcp dev file_converter_mcp_server.py
```

---

## Full Tool Reference

### PDF Core (11 tools)

| Tool | Description |
|------|-------------|
| `merge_pdfs` | Combine multiple PDFs into one |
| `split_pdf` | Split by page ranges or extract every page |
| `compress_pdf` | `screen` (72 DPI), `ebook` (150 DPI), `prepress` (lossless) |
| `rotate_pdf` | Rotate all or selected pages 90°, 180°, or 270° |
| `protect_pdf` | AES-256 password protection |
| `unlock_pdf` | Remove password from a protected PDF |
| `add_watermark` | Diagonal text or image watermark with opacity control |
| `add_page_numbers` | Customisable position, font size, starting number |
| `organize_pdf` | Reorder, duplicate, or delete pages |
| `repair_pdf` | Fix corrupted or malformed PDFs |
| `get_pdf_info` | Metadata, page count, dimensions, encryption status |

### PDF → Other (7 tools)

| Tool | Description |
|------|-------------|
| `pdf_to_images` | Render pages as PNG/JPG/WebP at any DPI |
| `pdf_to_text` | Extract plain text |
| `pdf_to_word` | Convert to editable `.docx` |
| `pdf_to_excel` | Extract tables to `.xlsx` |
| `pdf_to_pptx` | Convert to `.pptx` (image or text mode) |
| `pdf_to_html` | Convert to structured HTML |
| `pdf_to_markdown` | Convert to Markdown |

### Other → PDF (5 tools)

| Tool | Description |
|------|-------------|
| `images_to_pdf` | Combine images into a single PDF |
| `word_to_pdf` | `.docx` → PDF *(requires Microsoft Word on Windows)* |
| `excel_to_pdf` | `.xlsx` → PDF *(requires Microsoft Excel on Windows)* |
| `pptx_to_pdf` | `.pptx` → PDF *(requires Microsoft PowerPoint on Windows)* |
| `html_to_pdf` | HTML → PDF (WeasyPrint or ReportLab fallback) |

### OCR (6 tools)

| Tool | Description |
|------|-------------|
| `ocr_pdf` | OCR scanned PDF; optionally overlay invisible text layer |
| `ocr_image` | Extract text from any image (plain text or JSON with coords) |
| `image_to_word` | OCR image → `.docx` |
| `image_to_html` | OCR image → HTML |
| `image_to_markdown` | OCR image → Markdown |
| `list_ocr_languages` | List all 80+ supported EasyOCR language codes |

### Images (7 tools)

| Tool | Description |
|------|-------------|
| `convert_image` | Convert between PNG, JPG, WebP, BMP, GIF, TIFF with optional resize |
| `get_image_info` | Format, mode, dimensions, file size |
| `gif_to_frames` | Extract every frame from an animated GIF |
| `frames_to_gif` | Combine images into an animated GIF |
| `image_to_base64` | Encode an image as a Base64 data URI |
| `base64_to_image` | Decode a Base64 string back to an image |
| `heic_to_image` | Convert Apple HEIC/HEIF photos to JPG, PNG, or WebP |

### Documents (20 tools)

| Input | Outputs |
|-------|--------|
| `.docx` | HTML, Markdown, TXT |
| `.html` | DOCX, Markdown, TXT, XLSX |
| `.md` | HTML, PDF, DOCX, TXT |
| `.txt` | PDF, DOCX, HTML |
| `.pptx` | TXT, HTML, Images |
| `.xlsx` | HTML, Markdown, DOCX |

### Data (18 tools)

| Tool | Description |
|------|-------------|
| `convert_data` | JSON ↔ YAML ↔ CSV ↔ TOML ↔ XLSX |
| `xml_to_json` / `json_to_xml` | XML ↔ JSON |
| `xml_to_csv` | Flatten XML to CSV |
| `ini_to_json` / `ini_to_yaml` | INI config → JSON / YAML |
| `env_to_json` | `.env` file → JSON |
| `csv_to_markdown` | CSV → Markdown table |
| `html_table_to_csv` | Extract HTML `<table>` → CSV |
| `jsonl_to_json` / `json_to_jsonl` | JSONL ↔ JSON array |
| `csv_to_parquet` / `parquet_to_csv` | CSV ↔ Parquet |
| `parquet_to_json` | Parquet → JSON |
| `sqlite_to_csv` / `sqlite_to_json` / `sqlite_to_xlsx` | SQLite → CSV/JSON/XLSX |

### Archives (6 tools)

| Tool | Description |
|------|-------------|
| `zip_files` / `unzip_files` | Create and extract ZIP archives |
| `tar_files` / `untar_files` | Create and extract TAR archives |
| `create_7z` / `extract_7z` | Create and extract 7Z archives |

### E-books (7 tools)

| Tool | Description |
|------|-------------|
| `epub_to_txt/html/markdown/docx/pdf` | Convert EPUB to various formats |
| `rtf_to_txt` / `rtf_to_html` | Convert RTF documents |

---

## Architecture

```
file-converter-mcp/
├── file_converter_mcp_server.py   # MCP server entry point (stdio / SSE)
├── web_app.py                     # FastAPI web server + drag-and-drop UI
├── static/index.html              # FileForge web UI (single file, no build step)
├── tools/                         # One module per category
└── utils/                         # Shared helpers (file paths, PDF, OCR)
```

All tools return `{"success": True, ...}` or `{"success": False, "error": "..."}` — never raise exceptions.

---

## Limitations

| Area | Notes |
|------|-------|
| **PDF → Word/Excel fidelity** | Open-source can't match Adobe. Complex layouts need cleanup. |
| **Office → PDF** | Requires Microsoft Office on Windows (COM automation). |
| **OCR models** | First call downloads EasyOCR weights (~100–400 MB). |
| **HEIC** | Requires `pip install pillow-heif`. |
| **E-books** | Require `pip install ebooklib striprtf`. |
| **Audio/Video** | Not supported (would need ffmpeg). |

---

## Contributing

PRs welcome! High-value areas: Docker image, audio/video (ffmpeg), SVG support, better PDF→Word fidelity.

---

## License

MIT — see [LICENSE](LICENSE)

---

<p align="center"><strong>If this saved you time, please ⭐ the repo — it helps others find it!</strong></p>
