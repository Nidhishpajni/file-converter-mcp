"""PDF import tools: convert images, Word, Excel, PPTX, and HTML to PDF."""
import io
import html as html_module
from pathlib import Path
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP

from utils.file_utils import resolve_input, resolve_output, make_output_path


def register_pdf_import_tools(app: FastMCP) -> None:

    @app.tool()
    def images_to_pdf(
        input_paths: List[str],
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Combine one or more images (PNG, JPEG, TIFF, WebP, BMP) into a single PDF.

        Images are placed one per page in the order provided.
        JPEG images are embedded losslessly (no re-encoding).
        """
        try:
            if not input_paths:
                return {"success": False, "error": "Provide at least one image path."}

            sources = [resolve_input(p) for p in input_paths]
            if not output_path:
                output_path = make_output_path(str(sources[0]), "_combined", ".pdf")
            dst = resolve_output(output_path)

            import img2pdf
            from PIL import Image

            prepared = []
            for src in sources:
                suffix = src.suffix.lower()
                if suffix in (".jpg", ".jpeg"):
                    prepared.append(src.read_bytes())
                else:
                    img = Image.open(src)
                    if img.mode in ("RGBA", "P", "LA"):
                        img = img.convert("RGB")
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=95)
                    prepared.append(buf.getvalue())

            pdf_bytes = img2pdf.convert(prepared)
            dst.write_bytes(pdf_bytes)

            return {
                "success": True,
                "output_path": str(dst),
                "images_combined": len(sources),
                "file_size_bytes": dst.stat().st_size,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def word_to_pdf(
        input_path: str,
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Convert a Word document (.docx or .doc) to PDF.

        Requires Microsoft Word to be installed (uses Windows COM automation).
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".pdf")
            dst = resolve_output(output_path)

            from docx2pdf import convert
            convert(str(src), str(dst))

            return {
                "success": True,
                "output_path": str(dst),
                "file_size_bytes": dst.stat().st_size,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except ImportError:
            return {"success": False, "error": "docx2pdf is not installed. Run: pip install docx2pdf"}
        except Exception as e:
            msg = str(e)
            if "win32" in msg.lower() or "com" in msg.lower() or "dispatch" in msg.lower():
                return {"success": False, "error": "Microsoft Word must be installed on Windows to use word_to_pdf."}
            return {"success": False, "error": f"Conversion failed: {msg}"}

    @app.tool()
    def excel_to_pdf(
        input_path: str,
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Convert an Excel spreadsheet (.xlsx or .xls) to PDF.

        Requires Microsoft Excel to be installed (uses Windows COM automation).
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".pdf")
            dst = resolve_output(output_path)

            from docx2pdf import convert
            convert(str(src), str(dst))

            return {
                "success": True,
                "output_path": str(dst),
                "file_size_bytes": dst.stat().st_size,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except ImportError:
            return {"success": False, "error": "docx2pdf is not installed. Run: pip install docx2pdf"}
        except Exception as e:
            msg = str(e)
            if "win32" in msg.lower() or "com" in msg.lower() or "dispatch" in msg.lower():
                return {"success": False, "error": "Microsoft Excel must be installed on Windows to use excel_to_pdf."}
            return {"success": False, "error": f"Conversion failed: {msg}"}

    @app.tool()
    def pptx_to_pdf(
        input_path: str,
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Convert a PowerPoint presentation (.pptx or .ppt) to PDF.

        Requires Microsoft PowerPoint to be installed (uses Windows COM automation).
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".pdf")
            dst = resolve_output(output_path)

            from docx2pdf import convert
            convert(str(src), str(dst))

            return {
                "success": True,
                "output_path": str(dst),
                "file_size_bytes": dst.stat().st_size,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except ImportError:
            return {"success": False, "error": "docx2pdf is not installed. Run: pip install docx2pdf"}
        except Exception as e:
            msg = str(e)
            if "win32" in msg.lower() or "com" in msg.lower() or "dispatch" in msg.lower():
                return {"success": False, "error": "Microsoft PowerPoint must be installed on Windows to use pptx_to_pdf."}
            return {"success": False, "error": f"Conversion failed: {msg}"}

    @app.tool()
    def html_to_pdf(
        output_path: str,
        input_path: Optional[str] = None,
        html_string: Optional[str] = None,
    ) -> Dict:
        """
        Convert HTML to PDF.

        Provide either input_path (path to an .html file) or html_string (raw HTML).
        Tries weasyprint first (best quality); falls back to a reportlab renderer
        for basic HTML if weasyprint is not installed.
        """
        try:
            if not input_path and not html_string:
                return {"success": False, "error": "Provide either input_path or html_string."}

            dst = resolve_output(output_path)

            if input_path:
                src = resolve_input(input_path)
                html_content = src.read_text(encoding="utf-8")
            else:
                html_content = html_string

            try:
                from weasyprint import HTML
                HTML(string=html_content).write_pdf(str(dst))
                method = "weasyprint"
            except ImportError:
                _html_to_pdf_reportlab(html_content, str(dst))
                method = "reportlab (basic)"

            return {
                "success": True,
                "output_path": str(dst),
                "file_size_bytes": dst.stat().st_size,
                "renderer": method,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": f"HTML to PDF failed: {e}"}


def _html_to_pdf_reportlab(html_content: str, output_path: str) -> None:
    """Minimal HTML → PDF using reportlab (strips tags, preserves text)."""
    import re
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.units import inch

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=inch, leftMargin=inch,
                            topMargin=inch, bottomMargin=inch)
    styles = getSampleStyleSheet()
    story = []

    text = re.sub(r"<(br|p|div|h[1-6]|li)[^>]*>", "\n", html_content, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html_module.unescape(text)

    for line in text.split("\n"):
        line = line.strip()
        if line:
            story.append(Paragraph(html_module.escape(line), styles["Normal"]))
            story.append(Spacer(1, 6))

    doc.build(story)
