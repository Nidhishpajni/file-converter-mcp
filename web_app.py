#!/usr/bin/env python
"""
FileForge Web App — FastAPI frontend for file-converter-mcp tools.

Run:
    uvicorn web_app:app --reload --port 8080
Then open http://localhost:8080
"""
import io
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

# ---------------------------------------------------------------------------
# ToolRegistry — mimics FastMCP so existing register_*_tools() work unchanged
# ---------------------------------------------------------------------------

class ToolRegistry:
    """Drop-in replacement for FastMCP that captures registered tool functions."""

    def __init__(self):
        self._tools: Dict[str, Callable] = {}

    def tool(self, fn: Optional[Callable] = None, **kwargs):
        def decorator(f: Callable) -> Callable:
            self._tools[f.__name__] = f
            return f
        return decorator(fn) if callable(fn) else decorator

    def call(self, name: str, **kwargs) -> Any:
        if name not in self._tools:
            raise KeyError(f"Unknown tool: {name}")
        return self._tools[name](**kwargs)

    def list_tools(self) -> List[str]:
        return list(self._tools.keys())


# ---------------------------------------------------------------------------
# Register all tool modules
# ---------------------------------------------------------------------------

registry = ToolRegistry()

# Add project root to sys.path so utils/ is importable
import sys
sys.path.insert(0, str(Path(__file__).parent))

from tools import (
    register_pdf_core_tools,
    register_pdf_export_tools,
    register_pdf_import_tools,
    register_ocr_tools,
    register_image_tools,
    register_data_tools,
    register_document_tools,
    register_archive_tools,
    register_ebook_tools,
)

register_pdf_core_tools(registry)
register_pdf_export_tools(registry)
register_pdf_import_tools(registry)
register_ocr_tools(registry)
register_image_tools(registry)
register_data_tools(registry)
register_document_tools(registry)
register_archive_tools(registry)
register_ebook_tools(registry)

# ---------------------------------------------------------------------------
# Tool metadata (name, description, accepted file types, options)
# ---------------------------------------------------------------------------

TOOL_CATEGORIES = [
    {
        "id": "pdf",
        "name": "PDF Tools",
        "icon": "📄",
        "color": "#ef4444",
        "tools": [
            {"id": "merge_pdfs", "name": "Merge PDFs", "description": "Combine multiple PDFs into one", "accepts": [".pdf"], "multi": True, "options": []},
            {"id": "split_pdf", "name": "Split PDF", "description": "Split PDF into separate files", "accepts": [".pdf"], "options": [
                {"name": "split_every_page", "label": "Split every page", "type": "checkbox", "default": True},
            ]},
            {"id": "compress_pdf", "name": "Compress PDF", "description": "Reduce PDF file size", "accepts": [".pdf"], "options": [
                {"name": "quality_preset", "label": "Quality", "type": "select", "options": ["screen", "ebook", "prepress"], "default": "ebook"},
            ]},
            {"id": "rotate_pdf", "name": "Rotate PDF", "description": "Rotate PDF pages", "accepts": [".pdf"], "options": [
                {"name": "rotation", "label": "Degrees", "type": "select", "options": ["90", "180", "270"], "default": "90"},
                {"name": "pages", "label": "Pages (e.g. 1,3,5 or leave blank for all)", "type": "text", "default": ""},
            ]},
            {"id": "protect_pdf", "name": "Protect PDF", "description": "Password-protect a PDF", "accepts": [".pdf"], "options": [
                {"name": "user_password", "label": "Password", "type": "password", "default": ""},
            ]},
            {"id": "unlock_pdf", "name": "Unlock PDF", "description": "Remove password from PDF", "accepts": [".pdf"], "options": [
                {"name": "password", "label": "Current Password", "type": "password", "default": ""},
            ]},
            {"id": "add_watermark", "name": "Add Watermark", "description": "Add text watermark to PDF", "accepts": [".pdf"], "options": [
                {"name": "text", "label": "Watermark Text", "type": "text", "default": "CONFIDENTIAL"},
                {"name": "opacity", "label": "Opacity (0-1)", "type": "number", "default": "0.3"},
            ]},
            {"id": "add_page_numbers", "name": "Page Numbers", "description": "Add page numbers to PDF", "accepts": [".pdf"], "options": [
                {"name": "position", "label": "Position", "type": "select", "options": ["bottom-center", "bottom-right", "bottom-left", "top-center"], "default": "bottom-center"},
                {"name": "start_number", "label": "Start Number", "type": "number", "default": "1"},
            ]},
            {"id": "organize_pdf", "name": "Organize PDF", "description": "Reorder pages in a PDF", "accepts": [".pdf"], "options": [
                {"name": "page_order", "label": "Page order (e.g. 3,1,2)", "type": "text", "default": ""},
            ]},
            {"id": "repair_pdf", "name": "Repair PDF", "description": "Fix corrupted PDF files", "accepts": [".pdf"], "options": []},
            {"id": "get_pdf_info", "name": "PDF Info", "description": "Get metadata and page info", "accepts": [".pdf"], "options": [], "info_only": True},
        ],
    },
    {
        "id": "pdf_export",
        "name": "PDF → Other",
        "icon": "📤",
        "color": "#3b82f6",
        "tools": [
            {"id": "pdf_to_images", "name": "PDF → Images", "description": "Export each page as an image", "accepts": [".pdf"], "options": [
                {"name": "output_format", "label": "Format", "type": "select", "options": ["png", "jpg", "webp"], "default": "png"},
                {"name": "dpi", "label": "DPI", "type": "number", "default": "150"},
            ]},
            {"id": "pdf_to_word", "name": "PDF → Word", "description": "Convert PDF to .docx", "accepts": [".pdf"], "options": []},
            {"id": "pdf_to_excel", "name": "PDF → Excel", "description": "Extract tables to .xlsx", "accepts": [".pdf"], "options": []},
            {"id": "pdf_to_pptx", "name": "PDF → PowerPoint", "description": "Convert PDF to .pptx", "accepts": [".pdf"], "options": []},
            {"id": "pdf_to_html", "name": "PDF → HTML", "description": "Convert PDF to HTML", "accepts": [".pdf"], "options": []},
            {"id": "pdf_to_text", "name": "PDF → Text", "description": "Extract plain text from PDF", "accepts": [".pdf"], "options": []},
            {"id": "pdf_to_markdown", "name": "PDF → Markdown", "description": "Convert PDF to Markdown", "accepts": [".pdf"], "options": []},
        ],
    },
    {
        "id": "pdf_import",
        "name": "Other → PDF",
        "icon": "📥",
        "color": "#22c55e",
        "tools": [
            {"id": "images_to_pdf", "name": "Images → PDF", "description": "Combine images into a PDF", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"], "multi": True, "options": []},
            {"id": "word_to_pdf", "name": "Word → PDF", "description": "Convert .docx to PDF", "accepts": [".docx", ".doc"], "options": []},
            {"id": "excel_to_pdf", "name": "Excel → PDF", "description": "Convert .xlsx to PDF", "accepts": [".xlsx", ".xls"], "options": []},
            {"id": "pptx_to_pdf", "name": "PowerPoint → PDF", "description": "Convert .pptx to PDF", "accepts": [".pptx", ".ppt"], "options": []},
            {"id": "html_to_pdf", "name": "HTML → PDF", "description": "Convert HTML file to PDF", "accepts": [".html", ".htm"], "options": []},
        ],
    },
    {
        "id": "ocr",
        "name": "OCR",
        "icon": "🔍",
        "color": "#a855f7",
        "tools": [
            {"id": "ocr_pdf", "name": "OCR PDF", "description": "Make scanned PDFs searchable", "accepts": [".pdf"], "options": [
                {"name": "create_searchable_pdf", "label": "Create searchable PDF", "type": "checkbox", "default": True},
                {"name": "dpi", "label": "DPI", "type": "number", "default": "200"},
            ]},
            {"id": "ocr_image", "name": "Image → Text (OCR)", "description": "Extract text from image", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"], "options": [
                {"name": "output_format", "label": "Output", "type": "select", "options": ["text", "json"], "default": "text"},
            ]},
            {"id": "image_to_word", "name": "Image → Word", "description": "OCR image to .docx", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp"], "options": []},
            {"id": "image_to_html", "name": "Image → HTML", "description": "OCR image to HTML", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp"], "options": []},
            {"id": "image_to_markdown", "name": "Image → Markdown", "description": "OCR image to Markdown", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp"], "options": []},
        ],
    },
    {
        "id": "image",
        "name": "Images",
        "icon": "🖼️",
        "color": "#14b8a6",
        "tools": [
            {"id": "convert_image", "name": "Convert Image", "description": "Convert between image formats", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff", ".tif", ".heic", ".heif"], "options": [
                {"name": "output_format", "label": "Output Format", "type": "select", "options": ["png", "jpg", "webp", "bmp", "gif", "tiff"], "default": "png"},
                {"name": "quality", "label": "Quality (1-95, JPEG/WebP)", "type": "number", "default": "85"},
                {"name": "resize_width", "label": "Width px (optional)", "type": "number", "default": ""},
                {"name": "resize_height", "label": "Height px (optional)", "type": "number", "default": ""},
            ]},
            {"id": "get_image_info", "name": "Image Info", "description": "Get image dimensions and metadata", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff"], "options": [], "info_only": True},
            {"id": "gif_to_frames", "name": "GIF → Frames", "description": "Extract frames from animated GIF", "accepts": [".gif"], "options": [
                {"name": "output_format", "label": "Frame Format", "type": "select", "options": ["png", "jpg"], "default": "png"},
            ]},
            {"id": "frames_to_gif", "name": "Frames → GIF", "description": "Create animated GIF from images", "accepts": [".jpg", ".jpeg", ".png"], "multi": True, "options": [
                {"name": "duration_ms", "label": "Duration per frame (ms)", "type": "number", "default": "100"},
            ]},
            {"id": "image_to_base64", "name": "Image → Base64", "description": "Encode image as Base64 string", "accepts": [".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"], "options": []},
            {"id": "heic_to_image", "name": "HEIC → Image", "description": "Convert Apple HEIC to JPG/PNG", "accepts": [".heic", ".heif"], "options": [
                {"name": "output_format", "label": "Output Format", "type": "select", "options": ["jpg", "png", "webp"], "default": "jpg"},
            ]},
        ],
    },
    {
        "id": "documents",
        "name": "Documents",
        "icon": "📝",
        "color": "#6366f1",
        "tools": [
            {"id": "docx_to_html", "name": "DOCX → HTML", "description": "Convert Word to HTML", "accepts": [".docx"], "options": []},
            {"id": "docx_to_markdown", "name": "DOCX → Markdown", "description": "Convert Word to Markdown", "accepts": [".docx"], "options": []},
            {"id": "docx_to_txt", "name": "DOCX → Text", "description": "Extract text from Word", "accepts": [".docx"], "options": []},
            {"id": "html_to_markdown", "name": "HTML → Markdown", "description": "Convert HTML to Markdown", "accepts": [".html", ".htm"], "options": []},
            {"id": "html_to_txt", "name": "HTML → Text", "description": "Extract text from HTML", "accepts": [".html", ".htm"], "options": []},
            {"id": "html_to_docx", "name": "HTML → DOCX", "description": "Convert HTML to Word", "accepts": [".html", ".htm"], "options": []},
            {"id": "html_to_xlsx", "name": "HTML → Excel", "description": "Extract tables to Excel", "accepts": [".html", ".htm"], "options": []},
            {"id": "markdown_to_html", "name": "Markdown → HTML", "description": "Convert Markdown to HTML", "accepts": [".md", ".markdown"], "options": []},
            {"id": "markdown_to_pdf", "name": "Markdown → PDF", "description": "Convert Markdown to PDF", "accepts": [".md", ".markdown"], "options": []},
            {"id": "markdown_to_docx", "name": "Markdown → DOCX", "description": "Convert Markdown to Word", "accepts": [".md", ".markdown"], "options": []},
            {"id": "txt_to_pdf", "name": "Text → PDF", "description": "Convert plain text to PDF", "accepts": [".txt"], "options": []},
            {"id": "txt_to_docx", "name": "Text → DOCX", "description": "Convert plain text to Word", "accepts": [".txt"], "options": []},
            {"id": "txt_to_html", "name": "Text → HTML", "description": "Convert plain text to HTML", "accepts": [".txt"], "options": []},
            {"id": "pptx_to_txt", "name": "PPTX → Text", "description": "Extract text from PowerPoint", "accepts": [".pptx"], "options": []},
            {"id": "pptx_to_html", "name": "PPTX → HTML", "description": "Convert PowerPoint to HTML", "accepts": [".pptx"], "options": []},
            {"id": "pptx_to_images", "name": "PPTX → Images", "description": "Export slides as images", "accepts": [".pptx"], "options": []},
            {"id": "xlsx_to_html", "name": "Excel → HTML", "description": "Convert spreadsheet to HTML", "accepts": [".xlsx"], "options": []},
            {"id": "xlsx_to_markdown", "name": "Excel → Markdown", "description": "Convert spreadsheet to Markdown", "accepts": [".xlsx"], "options": []},
            {"id": "xlsx_to_docx", "name": "Excel → DOCX", "description": "Convert spreadsheet to Word", "accepts": [".xlsx"], "options": []},
        ],
    },
    {
        "id": "data",
        "name": "Data",
        "icon": "📊",
        "color": "#f59e0b",
        "tools": [
            {"id": "convert_data", "name": "Data Convert", "description": "JSON / YAML / CSV / TOML / XLSX", "accepts": [".json", ".yaml", ".yml", ".csv", ".toml", ".xlsx"], "options": [
                {"name": "output_format", "label": "Output Format", "type": "select", "options": ["json", "yaml", "csv", "toml", "xlsx"], "default": "json"},
            ]},
            {"id": "xml_to_json", "name": "XML → JSON", "description": "Convert XML to JSON", "accepts": [".xml"], "options": []},
            {"id": "json_to_xml", "name": "JSON → XML", "description": "Convert JSON to XML", "accepts": [".json"], "options": []},
            {"id": "xml_to_csv", "name": "XML → CSV", "description": "Convert XML to CSV", "accepts": [".xml"], "options": []},
            {"id": "ini_to_json", "name": "INI → JSON", "description": "Convert config file to JSON", "accepts": [".ini", ".cfg"], "options": []},
            {"id": "ini_to_yaml", "name": "INI → YAML", "description": "Convert config file to YAML", "accepts": [".ini", ".cfg"], "options": []},
            {"id": "env_to_json", "name": "ENV → JSON", "description": "Convert .env file to JSON", "accepts": [".env"], "options": []},
            {"id": "csv_to_markdown", "name": "CSV → Markdown", "description": "Convert CSV to Markdown table", "accepts": [".csv"], "options": []},
            {"id": "html_table_to_csv", "name": "HTML Table → CSV", "description": "Extract HTML tables to CSV", "accepts": [".html", ".htm"], "options": []},
            {"id": "jsonl_to_json", "name": "JSONL → JSON", "description": "Convert JSON Lines to JSON array", "accepts": [".jsonl"], "options": []},
            {"id": "json_to_jsonl", "name": "JSON → JSONL", "description": "Convert JSON array to JSON Lines", "accepts": [".json"], "options": []},
            {"id": "sqlite_to_csv", "name": "SQLite → CSV", "description": "Export SQLite tables to CSV", "accepts": [".db", ".sqlite", ".sqlite3"], "options": []},
            {"id": "sqlite_to_json", "name": "SQLite → JSON", "description": "Export SQLite tables to JSON", "accepts": [".db", ".sqlite", ".sqlite3"], "options": []},
            {"id": "sqlite_to_xlsx", "name": "SQLite → Excel", "description": "Export SQLite tables to Excel", "accepts": [".db", ".sqlite", ".sqlite3"], "options": []},
            {"id": "parquet_to_csv", "name": "Parquet → CSV", "description": "Convert Parquet to CSV", "accepts": [".parquet"], "options": []},
            {"id": "parquet_to_json", "name": "Parquet → JSON", "description": "Convert Parquet to JSON", "accepts": [".parquet"], "options": []},
            {"id": "csv_to_parquet", "name": "CSV → Parquet", "description": "Convert CSV to Parquet", "accepts": [".csv"], "options": []},
        ],
    },
    {
        "id": "archives",
        "name": "Archives",
        "icon": "🗃️",
        "color": "#64748b",
        "tools": [
            {"id": "zip_files", "name": "Create ZIP", "description": "Compress files into a ZIP", "accepts": ["*"], "multi": True, "options": [
                {"name": "archive_name", "label": "Archive name (without .zip)", "type": "text", "default": "archive"},
            ]},
            {"id": "unzip_files", "name": "Extract ZIP", "description": "Extract files from a ZIP", "accepts": [".zip"], "options": []},
            {"id": "tar_files", "name": "Create TAR", "description": "Compress files into a TAR", "accepts": ["*"], "multi": True, "options": [
                {"name": "compression", "label": "Compression", "type": "select", "options": ["gz", "bz2", "xz", "none"], "default": "gz"},
            ]},
            {"id": "untar_files", "name": "Extract TAR", "description": "Extract files from a TAR", "accepts": [".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tar.xz"], "options": []},
            {"id": "create_7z", "name": "Create 7Z", "description": "Compress files into a 7Z archive", "accepts": ["*"], "multi": True, "options": []},
            {"id": "extract_7z", "name": "Extract 7Z", "description": "Extract files from a 7Z archive", "accepts": [".7z"], "options": []},
        ],
    },
    {
        "id": "ebooks",
        "name": "E-books",
        "icon": "📚",
        "color": "#ec4899",
        "tools": [
            {"id": "epub_to_txt", "name": "EPUB → Text", "description": "Extract text from EPUB", "accepts": [".epub"], "options": []},
            {"id": "epub_to_html", "name": "EPUB → HTML", "description": "Convert EPUB to HTML", "accepts": [".epub"], "options": []},
            {"id": "epub_to_markdown", "name": "EPUB → Markdown", "description": "Convert EPUB to Markdown", "accepts": [".epub"], "options": []},
            {"id": "epub_to_docx", "name": "EPUB → DOCX", "description": "Convert EPUB to Word", "accepts": [".epub"], "options": []},
            {"id": "epub_to_pdf", "name": "EPUB → PDF", "description": "Convert EPUB to PDF", "accepts": [".epub"], "options": []},
            {"id": "rtf_to_txt", "name": "RTF → Text", "description": "Extract text from RTF", "accepts": [".rtf"], "options": []},
            {"id": "rtf_to_html", "name": "RTF → HTML", "description": "Convert RTF to HTML", "accepts": [".rtf"], "options": []},
        ],
    },
]

# Output extension map for single-output tools
OUTPUT_EXT_MAP = {
    "merge_pdfs": ".pdf", "split_pdf": None, "compress_pdf": ".pdf",
    "rotate_pdf": ".pdf", "protect_pdf": ".pdf", "unlock_pdf": ".pdf",
    "add_watermark": ".pdf", "add_page_numbers": ".pdf", "organize_pdf": ".pdf",
    "repair_pdf": ".pdf", "get_pdf_info": None,
    "pdf_to_images": None, "pdf_to_word": ".docx", "pdf_to_excel": ".xlsx",
    "pdf_to_pptx": ".pptx", "pdf_to_html": ".html", "pdf_to_text": ".txt",
    "pdf_to_markdown": ".md",
    "images_to_pdf": ".pdf", "word_to_pdf": ".pdf", "excel_to_pdf": ".pdf",
    "pptx_to_pdf": ".pdf", "html_to_pdf": ".pdf",
    "ocr_pdf": ".pdf", "ocr_image": ".txt", "image_to_word": ".docx",
    "image_to_html": ".html", "image_to_markdown": ".md",
    "convert_image": None,  # determined by output_format option
    "get_image_info": None,
    "gif_to_frames": None, "frames_to_gif": ".gif",
    "image_to_base64": ".txt", "heic_to_image": None,
    "docx_to_html": ".html", "docx_to_markdown": ".md", "docx_to_txt": ".txt",
    "html_to_markdown": ".md", "html_to_txt": ".txt", "html_to_docx": ".docx",
    "html_to_xlsx": ".xlsx",
    "markdown_to_html": ".html", "markdown_to_pdf": ".pdf",
    "markdown_to_docx": ".docx", "markdown_to_txt": ".txt",
    "txt_to_pdf": ".pdf", "txt_to_docx": ".docx", "txt_to_html": ".html",
    "pptx_to_txt": ".txt", "pptx_to_html": ".html", "pptx_to_images": None,
    "xlsx_to_html": ".html", "xlsx_to_markdown": ".md", "xlsx_to_docx": ".docx",
    "convert_data": None,  # determined by output_format option
    "get_supported_formats": None,
    "xml_to_json": ".json", "json_to_xml": ".xml", "xml_to_csv": ".csv",
    "ini_to_json": ".json", "ini_to_yaml": ".yaml", "env_to_json": ".json",
    "csv_to_markdown": ".md", "html_table_to_csv": ".csv",
    "jsonl_to_json": ".json", "json_to_jsonl": ".jsonl",
    "sqlite_to_csv": None, "sqlite_to_json": ".json", "sqlite_to_xlsx": ".xlsx",
    "parquet_to_csv": ".csv", "parquet_to_json": ".json", "csv_to_parquet": ".parquet",
    "zip_files": ".zip", "unzip_files": None, "tar_files": ".tar.gz",
    "untar_files": None, "create_7z": ".7z", "extract_7z": None,
    "epub_to_txt": ".txt", "epub_to_html": ".html", "epub_to_markdown": ".md",
    "epub_to_docx": ".docx", "epub_to_pdf": ".pdf",
    "rtf_to_txt": ".txt", "rtf_to_html": ".html",
}

# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------

app = FastAPI(title="FileForge", description="Comprehensive file converter")


@app.get("/api/tools")
def get_tools():
    return {"categories": TOOL_CATEGORIES}


@app.post("/api/convert")
async def convert(
    tool: str = Form(...),
    options: str = Form("{}"),
    files: List[UploadFile] = File(...),
):
    opts = json.loads(options)
    tmpdir = tempfile.mkdtemp(prefix="fileforge_")

    try:
        # Save uploaded files
        saved_paths: List[str] = []
        for f in files:
            dest = os.path.join(tmpdir, f.filename or "upload")
            with open(dest, "wb") as fh:
                shutil.copyfileobj(f.file, fh)
            saved_paths.append(dest)

        primary = saved_paths[0] if saved_paths else ""

        # Build kwargs
        kwargs: Dict[str, Any] = {}

        # Multi-file tools
        multi_tools = {"merge_pdfs", "images_to_pdf", "frames_to_gif", "zip_files", "tar_files", "create_7z"}
        if tool in multi_tools:
            if tool == "merge_pdfs":
                kwargs["input_paths"] = saved_paths
            elif tool == "images_to_pdf":
                kwargs["input_paths"] = saved_paths
            elif tool == "frames_to_gif":
                kwargs["input_paths"] = saved_paths
                kwargs["output_path"] = os.path.join(tmpdir, "output.gif")
            elif tool in {"zip_files", "tar_files", "create_7z"}:
                kwargs["input_paths"] = saved_paths
        else:
            # Some tools use "file_path" instead of "input_path"
            FILE_PATH_TOOLS = {"get_image_info"}
            if tool in FILE_PATH_TOOLS:
                kwargs["file_path"] = primary
            else:
                kwargs["input_path"] = primary

        # Determine output path
        out_ext = OUTPUT_EXT_MAP.get(tool, ".out")

        # Dynamic ext based on options
        if tool == "convert_image" and opts.get("output_format"):
            out_ext = "." + opts["output_format"].lstrip(".")
        elif tool == "convert_data" and opts.get("output_format"):
            fmt = opts["output_format"]
            out_ext = {"json": ".json", "yaml": ".yaml", "csv": ".csv", "toml": ".toml", "xlsx": ".xlsx"}.get(fmt, ".json")
        elif tool == "heic_to_image" and opts.get("output_format"):
            out_ext = "." + opts["output_format"].lstrip(".")
        elif tool == "tar_files":
            comp = opts.get("compression", "gz")
            out_ext = ".tar" if comp == "none" else f".tar.{comp}"

        # Apply numeric coercions and page-list parsing
        for k, v in list(opts.items()):
            if v == "" or v is None:
                continue
            if k in ("dpi", "start_number", "duration_ms", "loop",
                     "quality", "resize_width", "resize_height", "font_size"):
                try:
                    opts[k] = int(float(v))
                except (ValueError, TypeError):
                    pass
            elif k == "opacity":
                try:
                    opts[k] = float(v)
                except (ValueError, TypeError):
                    pass
            elif k == "pages" and isinstance(v, str) and v.strip():
                try:
                    opts[k] = [int(x.strip()) for x in v.split(",") if x.strip()]
                except ValueError:
                    pass
            elif k == "page_order" and isinstance(v, str) and v.strip():
                try:
                    opts[k] = [int(x.strip()) for x in v.split(",") if x.strip()]
                except ValueError:
                    pass
            elif k == "rotation":
                try:
                    opts[k] = int(v)
                except (ValueError, TypeError):
                    pass
            elif k == "split_every_page":
                opts[k] = v in (True, "true", "on", "1", 1)
            elif k == "create_searchable_pdf":
                opts[k] = v in (True, "true", "on", "1", 1)

        # Set output path for single-output tools
        info_only_tools = {"get_pdf_info", "get_image_info", "get_supported_formats", "list_ocr_languages"}
        if tool not in info_only_tools and "output_path" not in kwargs:
            if out_ext:
                stem = Path(primary).stem if primary else "output"
                kwargs["output_path"] = os.path.join(tmpdir, f"{stem}_converted{out_ext}")

        # Merge user options
        for k, v in opts.items():
            if v not in ("", None):
                kwargs[k] = v

        # Archive name handling
        if tool == "zip_files":
            name = opts.get("archive_name", "archive")
            kwargs["output_path"] = os.path.join(tmpdir, f"{name}.zip")
            kwargs.pop("archive_name", None)
        elif tool == "create_7z":
            kwargs["output_path"] = os.path.join(tmpdir, "archive.7z")
        elif tool == "tar_files":
            comp = opts.get("compression", "gz")
            ext = ".tar" if comp == "none" else f".tar.{comp}"
            kwargs["output_path"] = os.path.join(tmpdir, f"archive{ext}")
            kwargs.pop("archive_name", None)

        # Info-only: call and return JSON
        if tool in info_only_tools:
            result = registry.call(tool, **kwargs)
            return JSONResponse(content=result)

        # Call the tool
        result = registry.call(tool, **kwargs)

        if not result.get("success"):
            return JSONResponse(status_code=400, content={"error": result.get("error", "Conversion failed")})

        # ---------- Handle multi-output tools ----------
        output_files_list = result.get("output_files")
        output_dir_val = result.get("output_dir")

        if output_files_list or (output_dir_val and not result.get("output_path")):
            # Zip all outputs
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                if output_files_list:
                    for fp in output_files_list:
                        zf.write(fp, Path(fp).name)
                elif output_dir_val:
                    for fp in Path(output_dir_val).iterdir():
                        if fp.is_file():
                            zf.write(str(fp), fp.name)
            zip_buf.seek(0)
            stem = Path(primary).stem if primary else "output"
            headers = {"Content-Disposition": f'attachment; filename="{stem}_output.zip"'}
            return StreamingResponse(zip_buf, media_type="application/zip", headers=headers,
                                     background=_cleanup(tmpdir))

        # ---------- Single output file ----------
        out_path = result.get("output_path")
        if not out_path or not Path(out_path).exists():
            return JSONResponse(status_code=400, content={"error": "Tool ran but produced no output file."})

        out_path = Path(out_path)
        media_type = _guess_media_type(out_path.suffix)
        headers = {"Content-Disposition": f'attachment; filename="{out_path.name}"'}

        # Stream and cleanup
        def iter_file():
            with open(out_path, "rb") as fh:
                while chunk := fh.read(65536):
                    yield chunk
            shutil.rmtree(tmpdir, ignore_errors=True)

        return StreamingResponse(iter_file(), media_type=media_type, headers=headers)

    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        return JSONResponse(status_code=500, content={"error": str(e)})


def _cleanup(tmpdir: str):
    """Background task to remove temp dir."""
    from starlette.background import BackgroundTask
    return BackgroundTask(shutil.rmtree, tmpdir, ignore_errors=True)


def _guess_media_type(ext: str) -> str:
    return {
        ".pdf": "application/pdf",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".html": "text/html",
        ".txt": "text/plain",
        ".md": "text/markdown",
        ".json": "application/octet-stream",
        ".yaml": "text/yaml",
        ".csv": "text/csv",
        ".xml": "application/xml",
        ".zip": "application/zip",
        ".7z": "application/x-7z-compressed",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
        ".bmp": "image/bmp",
        ".tiff": "image/tiff",
        ".parquet": "application/octet-stream",
        ".jsonl": "application/jsonl",
        ".toml": "application/toml",
        ".epub": "application/epub+zip",
    }.get(ext.lower(), "application/octet-stream")


# ---------------------------------------------------------------------------
# Static files (must come last so API routes take priority)
# ---------------------------------------------------------------------------

static_dir = Path(__file__).parent / "static"
static_dir.mkdir(exist_ok=True)
app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8080, reload=True)
