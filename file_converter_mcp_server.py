#!/usr/bin/env python
"""
file-converter-mcp — ilovepdf-style file conversion MCP server.

Exposes 83 tools across 9 categories:
  PDF core/export/import, OCR, Images, Documents, Data, Archives, E-books.

Usage:
  python file_converter_mcp_server.py                         # stdio (for MCP clients)
  python file_converter_mcp_server.py --transport sse --port 8000
"""
import argparse
from typing import Dict

from mcp.server.fastmcp import FastMCP

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

app = FastMCP(
    name="file-converter-mcp",
    instructions="A comprehensive file conversion server. Call get_server_info to see all available tools.",
)

register_pdf_core_tools(app)
register_pdf_export_tools(app)
register_pdf_import_tools(app)
register_ocr_tools(app)
register_image_tools(app)
register_data_tools(app)
register_document_tools(app)
register_archive_tools(app)
register_ebook_tools(app)


@app.tool()
def get_server_info() -> Dict:
    """Return information about this server and all available tools."""
    optional_status = {}
    for lib, label in [("weasyprint", "weasyprint"), ("docx2pdf", "docx2pdf"), ("easyocr", "easyocr")]:
        try:
            __import__(lib)
            optional_status[label] = "available"
        except ImportError:
            optional_status[label] = "not installed"
    try:
        import fitz
        optional_status["pymupdf"] = fitz.version[0]
    except ImportError:
        optional_status["pymupdf"] = "not installed"

    return {
        "name": "file-converter-mcp",
        "version": "2.0.0",
        "description": "Comprehensive file conversion MCP server (ilovepdf-style + more)",
        "total_tools": 83,
        "categories": [
            "pdf_core (11)", "pdf_export (7)", "pdf_import (5)",
            "ocr (6)", "image (7)", "documents (20)",
            "data (18)", "archives (6)", "ebooks (7)",
        ],
        "optional_dependencies": optional_status,
        "notes": {
            "ocr": "First OCR call downloads EasyOCR model weights (~100-400 MB to ~/.EasyOCR/).",
            "office_to_pdf": "word_to_pdf / excel_to_pdf / pptx_to_pdf require Microsoft Office on Windows.",
            "pdf_fidelity": "pdf_to_word/excel best on text-layer PDFs; use ocr_pdf first on scanned docs.",
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="file-converter-mcp MCP Server")
    parser.add_argument("--transport", choices=["stdio", "sse"], default="stdio")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()
    if args.transport == "sse":
        app.run(transport="sse", port=args.port)
    else:
        app.run(transport="stdio")


if __name__ == "__main__":
    main()
