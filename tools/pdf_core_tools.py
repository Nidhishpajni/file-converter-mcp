"""PDF core manipulation tools: merge, split, compress, rotate, protect, unlock,
watermark, page numbers, organize, repair, and get_pdf_info."""
import os
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF
from mcp.server.fastmcp import FastMCP

from utils.file_utils import resolve_input, resolve_output, make_output_path
from utils.pdf_utils import open_pdf, save_pdf


def register_pdf_core_tools(app: FastMCP) -> None:

    @app.tool()
    def merge_pdfs(
        input_paths: List[str],
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Merge multiple PDF files into a single PDF.

        input_paths: ordered list of PDF file paths to merge.
        output_path: destination file (default: first file's dir, named 'merged.pdf').
        """
        try:
            if len(input_paths) < 2:
                return {"success": False, "error": "Provide at least two input PDFs."}

            sources = [resolve_input(p) for p in input_paths]
            if not output_path:
                output_path = str(sources[0].parent / "merged.pdf")
            dst = resolve_output(output_path)

            merged = fitz.open()
            total_pages = 0
            for src in sources:
                doc = open_pdf(str(src))
                merged.insert_pdf(doc)
                total_pages += doc.page_count
                doc.close()

            merged.save(str(dst), garbage=3, deflate=True)
            merged.close()

            return {
                "success": True,
                "output_path": str(dst),
                "total_pages": total_pages,
                "files_merged": len(sources),
                "file_size_bytes": dst.stat().st_size,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def split_pdf(
        input_path: str,
        output_dir: Optional[str] = None,
        page_ranges: Optional[List[str]] = None,
        split_every_page: bool = False,
    ) -> Dict:
        """
        Split a PDF into multiple files.

        page_ranges: list of range strings like ["1-3", "4-6", "7"] (1-based).
        split_every_page: if True, extract each page as a separate file.
        output_dir: directory for output files (default: same dir as input).
        """
        try:
            src = resolve_input(input_path)
            doc = open_pdf(str(src))
            n = doc.page_count

            out_dir = Path(output_dir) if output_dir else src.parent
            out_dir.mkdir(parents=True, exist_ok=True)
            stem = src.stem

            output_files = []

            def extract_pages(indices: List[int], filename: str):
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=indices[0], to_page=indices[-1])
                path = str(out_dir / filename)
                new_doc.save(path, garbage=3)
                new_doc.close()
                return path

            if split_every_page:
                for i in range(n):
                    path = extract_pages([i], f"{stem}_page{i+1}.pdf")
                    output_files.append(path)
            elif page_ranges:
                for rng in page_ranges:
                    rng = rng.strip()
                    if "-" in rng:
                        a, b = rng.split("-", 1)
                        start, end = int(a) - 1, int(b) - 1
                    else:
                        start = end = int(rng) - 1
                    start = max(0, min(start, n - 1))
                    end = max(start, min(end, n - 1))
                    fname = f"{stem}_p{start+1}-{end+1}.pdf"
                    path = extract_pages(list(range(start, end + 1)), fname)
                    output_files.append(path)
            else:
                mid = n // 2
                output_files.append(extract_pages(list(range(0, mid)), f"{stem}_part1.pdf"))
                output_files.append(extract_pages(list(range(mid, n)), f"{stem}_part2.pdf"))

            doc.close()
            return {
                "success": True,
                "output_files": output_files,
                "files_created": len(output_files),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def compress_pdf(
        input_path: str,
        output_path: Optional[str] = None,
        quality_preset: str = "ebook",
    ) -> Dict:
        """
        Compress a PDF to reduce file size.

        quality_preset:
          - 'screen'   : aggressive compression, images downsampled to 72 DPI
          - 'ebook'    : moderate compression, images at 150 DPI (default)
          - 'prepress' : lossless stream compression only, images untouched
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_compressed")
            dst = resolve_output(output_path)
            original_size = src.stat().st_size

            preset_dpi = {"screen": 72, "ebook": 150, "prepress": None}
            if quality_preset not in preset_dpi:
                return {"success": False, "error": f"Unknown preset '{quality_preset}'. Use screen, ebook, or prepress."}

            doc = open_pdf(str(src))

            target_dpi = preset_dpi[quality_preset]
            if target_dpi:
                from PIL import Image
                import io
                for page in doc:
                    for img in page.get_images(full=True):
                        xref = img[0]
                        base_image = doc.extract_image(xref)
                        img_bytes = base_image["image"]
                        pil_img = Image.open(io.BytesIO(img_bytes))
                        w, h = pil_img.size
                        scale = target_dpi / 96
                        new_w, new_h = max(1, int(w * scale)), max(1, int(h * scale))
                        if new_w < w:
                            pil_img = pil_img.resize((new_w, new_h), Image.LANCZOS)
                            if pil_img.mode == "RGBA":
                                pil_img = pil_img.convert("RGB")
                            buf = io.BytesIO()
                            pil_img.save(buf, format="JPEG", quality=75, optimize=True)
                            doc.update_stream(xref, buf.getvalue())

            doc.save(
                str(dst),
                garbage=4,
                deflate=True,
                clean=True,
                deflate_images=True,
                deflate_fonts=True,
            )
            doc.close()

            new_size = dst.stat().st_size
            reduction = round((1 - new_size / original_size) * 100, 1) if original_size else 0

            return {
                "success": True,
                "output_path": str(dst),
                "quality_preset": quality_preset,
                "original_size_bytes": original_size,
                "compressed_size_bytes": new_size,
                "size_reduction_percent": reduction,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def rotate_pdf(
        input_path: str,
        rotation: int,
        output_path: Optional[str] = None,
        pages: Optional[List[int]] = None,
    ) -> Dict:
        """
        Rotate pages in a PDF.

        rotation: degrees clockwise — 90, 180, or 270.
        pages: 1-based page numbers to rotate (default: all pages).
        """
        try:
            if rotation not in (90, 180, 270):
                return {"success": False, "error": "rotation must be 90, 180, or 270."}

            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, f"_rotated{rotation}")
            dst = resolve_output(output_path)

            doc = open_pdf(str(src))
            n = doc.page_count
            targets = [p - 1 for p in pages] if pages else list(range(n))
            targets = [i for i in targets if 0 <= i < n]

            for i in targets:
                page = doc[i]
                page.set_rotation((page.rotation + rotation) % 360)

            save_pdf(doc, str(dst), garbage=3, deflate=True)
            doc.close()

            return {
                "success": True,
                "output_path": str(dst),
                "pages_rotated": len(targets),
                "rotation_degrees": rotation,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def protect_pdf(
        input_path: str,
        user_password: str,
        output_path: Optional[str] = None,
        owner_password: Optional[str] = None,
        allow_printing: bool = True,
        allow_copying: bool = False,
    ) -> Dict:
        """
        Password-protect a PDF with AES-256 encryption.

        user_password: password required to open the file.
        owner_password: password for full permissions (defaults to user_password).
        allow_printing: whether opening users can print.
        allow_copying: whether opening users can copy text.
        """
        try:
            import pikepdf

            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_protected")
            dst = resolve_output(output_path)

            owner_pw = owner_password or user_password

            pdf = pikepdf.open(str(src))
            perms = pikepdf.Permissions(
                print_highres=allow_printing,
                print_lowres=allow_printing,
                extract=allow_copying,
                modify_annotation=False,
                modify_assembly=False,
                modify_form=False,
                modify_other=False,
            )
            pdf.save(
                str(dst),
                encryption=pikepdf.Encryption(
                    owner=owner_pw,
                    user=user_password,
                    R=6,
                    allow=perms,
                ),
            )
            pdf.close()

            return {
                "success": True,
                "output_path": str(dst),
                "encryption": "AES-256",
                "allow_printing": allow_printing,
                "allow_copying": allow_copying,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def unlock_pdf(
        input_path: str,
        password: str,
        output_path: Optional[str] = None,
    ) -> Dict:
        """Remove password protection from a PDF (requires the correct password)."""
        try:
            import pikepdf

            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_unlocked")
            dst = resolve_output(output_path)

            pdf = pikepdf.open(str(src), password=password)
            pdf.save(str(dst))
            pdf.close()

            return {"success": True, "output_path": str(dst)}
        except Exception as e:
            if "password" in str(e).lower() or "encrypted" in str(e).lower():
                return {"success": False, "error": "Incorrect password or PDF is not encrypted."}
            return {"success": False, "error": str(e)}

    @app.tool()
    def add_watermark(
        input_path: str,
        output_path: Optional[str] = None,
        text: Optional[str] = None,
        image_path: Optional[str] = None,
        opacity: float = 0.3,
        pages: Optional[List[int]] = None,
        font_size: int = 48,
        color: str = "gray",
    ) -> Dict:
        """
        Add a text or image watermark to PDF pages.

        text: watermark text (e.g. 'CONFIDENTIAL').
        image_path: path to a PNG/JPEG watermark image.
        opacity: 0.0 (invisible) to 1.0 (opaque), default 0.3.
        pages: 1-based page list (default: all pages).
        color: text color name (gray, red, blue, black).
        """
        try:
            if not text and not image_path:
                return {"success": False, "error": "Provide either 'text' or 'image_path'."}

            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_watermarked")
            dst = resolve_output(output_path)

            doc = open_pdf(str(src))
            n = doc.page_count
            targets = [p - 1 for p in pages] if pages else list(range(n))
            targets = [i for i in targets if 0 <= i < n]

            _color_map = {
                "gray": (0.5, 0.5, 0.5), "red": (1, 0, 0),
                "blue": (0, 0, 1), "black": (0, 0, 0),
            }
            rgb = _color_map.get(color.lower(), (0.5, 0.5, 0.5))

            for i in targets:
                page = doc[i]
                rect = page.rect
                cx, cy = rect.width / 2, rect.height / 2

                if text:
                    import math
                    tw = fitz.TextWriter(page.rect, color=rgb, opacity=opacity)
                    font = fitz.Font("helv")
                    tw.append(fitz.Point(0, 0), text, font=font, fontsize=font_size)
                    angle = 45 * math.pi / 180
                    cos_a, sin_a = math.cos(angle), math.sin(angle)
                    morph = (
                        fitz.Point(cx, cy),
                        fitz.Matrix(cos_a, -sin_a, sin_a, cos_a, 0, 0),
                    )
                    tw.write_text(page, opacity=opacity, morph=morph, overlay=True)
                elif image_path:
                    img_src = resolve_input(image_path)
                    wm_w = rect.width * 0.5
                    wm_h = rect.height * 0.5
                    wm_rect = fitz.Rect(cx - wm_w / 2, cy - wm_h / 2, cx + wm_w / 2, cy + wm_h / 2)
                    page.insert_image(wm_rect, filename=str(img_src), overlay=True)

            save_pdf(doc, str(dst), garbage=3, deflate=True)
            doc.close()

            return {
                "success": True,
                "output_path": str(dst),
                "pages_watermarked": len(targets),
                "watermark_type": "text" if text else "image",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def add_page_numbers(
        input_path: str,
        output_path: Optional[str] = None,
        position: str = "bottom-center",
        start_number: int = 1,
        font_size: int = 11,
        margin: int = 20,
    ) -> Dict:
        """
        Add page numbers to every page of a PDF.

        position: 'bottom-center', 'bottom-right', 'bottom-left',
                  'top-center', 'top-right', 'top-left'.
        start_number: number to assign to the first page (default 1).
        margin: distance from the page edge in points.
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_numbered")
            dst = resolve_output(output_path)

            doc = open_pdf(str(src))

            for i, page in enumerate(doc):
                rect = page.rect
                num = start_number + i
                text = str(num)

                pos = position.lower()
                if "bottom" in pos:
                    y = rect.height - margin
                else:
                    y = margin + font_size

                if "center" in pos:
                    x = rect.width / 2 - len(text) * font_size * 0.3
                elif "right" in pos:
                    x = rect.width - margin - len(text) * font_size * 0.6
                else:
                    x = margin

                page.insert_text(
                    fitz.Point(x, y),
                    text,
                    fontsize=font_size,
                    color=(0, 0, 0),
                    overlay=True,
                )

            save_pdf(doc, str(dst), garbage=3, deflate=True)
            doc.close()

            return {
                "success": True,
                "output_path": str(dst),
                "pages_numbered": i + 1,
                "start_number": start_number,
                "position": position,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def organize_pdf(
        input_path: str,
        page_order: List[int],
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Reorder, duplicate, or delete pages in a PDF.

        page_order: list of 1-based page numbers defining the new order.
                    Repeat a number to duplicate; omit a number to delete.
        Example: [3, 1, 2] reorders a 3-page PDF; [1, 1, 2] duplicates page 1.
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_organized")
            dst = resolve_output(output_path)

            doc = open_pdf(str(src))
            n = doc.page_count

            indices = [p - 1 for p in page_order]
            invalid = [p for p in indices if not (0 <= p < n)]
            if invalid:
                return {"success": False, "error": f"Page numbers out of range (PDF has {n} pages): {[i+1 for i in invalid]}"}

            new_doc = fitz.open()
            for idx in indices:
                new_doc.insert_pdf(doc, from_page=idx, to_page=idx)

            save_pdf(new_doc, str(dst), garbage=3, deflate=True)
            doc.close()
            new_doc.close()

            return {
                "success": True,
                "output_path": str(dst),
                "original_pages": n,
                "output_pages": len(indices),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def repair_pdf(
        input_path: str,
        output_path: Optional[str] = None,
    ) -> Dict:
        """
        Attempt to repair a corrupted or malformed PDF.

        Fixes cross-reference table issues and stream errors.
        Cannot recover from truly corrupted binary data.
        """
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "_repaired")
            dst = resolve_output(output_path)

            doc = fitz.open(str(src), filetype="pdf")
            doc.save(str(dst), garbage=4, deflate=True, clean=True)
            page_count = doc.page_count
            doc.close()

            return {
                "success": True,
                "output_path": str(dst),
                "page_count": page_count,
                "note": "Basic structural repair applied. Severely corrupted data may not be recoverable.",
            }
        except Exception as e:
            return {"success": False, "error": f"Repair failed: {e}"}

    @app.tool()
    def get_pdf_info(input_path: str, password: Optional[str] = None) -> Dict:
        """
        Return metadata and properties of a PDF file.

        Returns: page_count, page_sizes, title, author, creator, encrypted status,
                 file size, and PDF version.
        """
        try:
            src = resolve_input(input_path)
            doc = open_pdf(str(src), password=password)

            meta = doc.metadata or {}
            pages_info = []
            for i, page in enumerate(doc):
                r = page.rect
                pages_info.append({"page": i + 1, "width_pt": round(r.width, 2), "height_pt": round(r.height, 2)})

            doc.close()

            return {
                "success": True,
                "path": str(src),
                "file_size_bytes": src.stat().st_size,
                "page_count": len(pages_info),
                "pages": pages_info,
                "title": meta.get("title", ""),
                "author": meta.get("author", ""),
                "subject": meta.get("subject", ""),
                "creator": meta.get("creator", ""),
                "producer": meta.get("producer", ""),
                "creation_date": meta.get("creationDate", ""),
                "modification_date": meta.get("modDate", ""),
                "encrypted": False,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except ValueError as e:
            return {"success": False, "error": str(e), "encrypted": True}
        except Exception as e:
            return {"success": False, "error": str(e)}
