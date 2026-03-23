"""Image conversion and inspection tools."""
import os
from pathlib import Path
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP

from utils.file_utils import resolve_input, resolve_output, make_output_path

_FORMAT_MAP = {
    "jpg": "JPEG", "jpeg": "JPEG",
    "png": "PNG",
    "webp": "WEBP",
    "bmp": "BMP",
    "gif": "GIF",
    "tif": "TIFF", "tiff": "TIFF",
}

_SAVE_KWARGS = {
    "JPEG": {"optimize": True},
    "WEBP": {"method": 4},
    "PNG":  {"optimize": True},
    "TIFF": {"compression": "tiff_lzw"},
    "GIF":  {},
    "BMP":  {},
}

SUPPORTED_FORMATS = ["png", "jpg", "webp", "bmp", "gif", "tiff"]


def _normalize(fmt: str) -> Optional[str]:
    return _FORMAT_MAP.get(fmt.lower())


def register_image_tools(app: FastMCP) -> None:

    @app.tool()
    def convert_image(
        input_path: str,
        output_path: Optional[str] = None,
        output_format: Optional[str] = None,
        quality: int = 85,
        resize_width: Optional[int] = None,
        resize_height: Optional[int] = None,
    ) -> Dict:
        """
        Convert an image to a different format, with optional resize.

        Supported formats: png, jpg, webp, bmp, gif, tiff.
        quality applies to JPEG and WebP (1-95, default 85).
        If only resize_width or resize_height is given, the other is calculated
        to preserve the aspect ratio.
        """
        from PIL import Image

        try:
            src = resolve_input(input_path)

            # Determine output format
            if output_format:
                pil_fmt = _normalize(output_format)
                if not pil_fmt:
                    return {"success": False, "error": f"Unsupported output format: {output_format}. Choose from {SUPPORTED_FORMATS}"}
                ext = "." + (output_format.lower() if output_format.lower() != "jpeg" else "jpg")
            else:
                # Infer from output_path extension
                if not output_path:
                    return {"success": False, "error": "Provide output_format or an output_path with a recognised extension."}
                ext = Path(output_path).suffix.lower()
                pil_fmt = _normalize(ext.lstrip("."))
                if not pil_fmt:
                    return {"success": False, "error": f"Cannot infer format from extension '{ext}'. Specify output_format."}

            if not output_path:
                output_path = make_output_path(input_path, "_converted", ext)
            dst = resolve_output(output_path)

            img = Image.open(src)
            original_size = img.size

            # Resize
            if resize_width or resize_height:
                w, h = img.size
                if resize_width and resize_height:
                    new_size = (resize_width, resize_height)
                elif resize_width:
                    new_size = (resize_width, int(h * resize_width / w))
                else:
                    new_size = (int(w * resize_height / h), resize_height)
                from PIL import Image as PILImage
                img = img.resize(new_size, PILImage.LANCZOS)

            # Mode conversion for formats that don't support alpha
            if pil_fmt == "JPEG" and img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            elif pil_fmt == "GIF" and img.mode != "P":
                img = img.convert("P")
            elif pil_fmt not in ("PNG", "WEBP") and img.mode == "RGBA":
                img = img.convert("RGB")

            kwargs = dict(_SAVE_KWARGS.get(pil_fmt, {}))
            if pil_fmt in ("JPEG", "WEBP"):
                kwargs["quality"] = max(1, min(95, quality))

            img.save(str(dst), format=pil_fmt, **kwargs)

            return {
                "success": True,
                "input_path": str(src),
                "output_path": str(dst),
                "input_format": src.suffix.lstrip(".").upper(),
                "output_format": pil_fmt,
                "original_size": list(original_size),
                "output_size": list(img.size),
                "file_size_bytes": dst.stat().st_size,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": f"Image conversion failed: {e}"}

    @app.tool()
    def get_image_info(file_path: str) -> Dict:
        """Return metadata about an image file: format, mode, dimensions, file size."""
        from PIL import Image

        try:
            src = resolve_input(file_path)
            img = Image.open(src)
            return {
                "success": True,
                "path": str(src),
                "format": img.format or src.suffix.lstrip(".").upper(),
                "mode": img.mode,
                "width": img.size[0],
                "height": img.size[1],
                "file_size_bytes": src.stat().st_size,
            }
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": f"Could not read image: {e}"}

    # ── GIF → Frames ────────────────────────────────────────────────────────
    @app.tool()
    def gif_to_frames(
        input_path: str,
        output_dir: Optional[str] = None,
        output_format: str = "png",
    ) -> Dict:
        """Extract individual frames from an animated GIF as separate image files."""
        from PIL import Image
        try:
            src = resolve_input(input_path)
            out_dir = Path(output_dir) if output_dir else src.parent / (src.stem + "_frames")
            out_dir.mkdir(parents=True, exist_ok=True)

            img = Image.open(src)
            frames = []
            ext = f".{output_format.lower()}"
            try:
                for i in range(img.n_frames):
                    img.seek(i)
                    frame = img.convert("RGBA") if output_format.lower() == "png" else img.convert("RGB")
                    out_path = out_dir / f"frame_{i:04d}{ext}"
                    frame.save(str(out_path))
                    frames.append(str(out_path))
            except EOFError:
                pass
            return {"success": True, "output_dir": str(out_dir), "frames_extracted": len(frames), "output_files": frames}
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── Frames → Animated GIF ───────────────────────────────────────────────
    @app.tool()
    def frames_to_gif(
        input_paths: List[str],
        output_path: str,
        duration_ms: int = 100,
        loop: int = 0,
    ) -> Dict:
        """
        Combine image files into an animated GIF.

        input_paths: ordered list of image paths (PNG, JPG, etc.).
        duration_ms: milliseconds per frame (default 100).
        loop: number of loops, 0 = infinite.
        """
        from PIL import Image
        try:
            if not input_paths:
                return {"success": False, "error": "Provide at least one image."}
            dst = resolve_output(output_path)
            frames = [Image.open(p).convert("RGBA") for p in input_paths]
            frames[0].save(
                str(dst), format="GIF", save_all=True,
                append_images=frames[1:], duration=duration_ms, loop=loop
            )
            return {"success": True, "output_path": str(dst), "frames": len(frames), "file_size_bytes": dst.stat().st_size}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── Image ↔ Base64 ──────────────────────────────────────────────────────
    @app.tool()
    def image_to_base64(input_path: str, output_path: Optional[str] = None, include_data_uri: bool = True) -> Dict:
        """
        Encode an image as a Base64 string.

        include_data_uri: if True, wraps in a data URI (data:image/png;base64,...).
        If output_path is omitted and image is small (<500 KB), returns base64 inline.
        """
        import base64
        try:
            src = resolve_input(input_path)
            raw = src.read_bytes()
            b64 = base64.b64encode(raw).decode("ascii")
            mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                    "gif": "image/gif", "webp": "image/webp", "bmp": "image/bmp",
                    "tiff": "image/tiff", "svg": "image/svg+xml"}.get(src.suffix.lstrip(".").lower(), "image/png")
            result = f"data:{mime};base64,{b64}" if include_data_uri else b64

            if output_path:
                dst = resolve_output(output_path)
                dst.write_text(result, encoding="ascii")
                return {"success": True, "output_path": str(dst), "mime_type": mime, "size_chars": len(result)}
            elif len(raw) < 512_000:
                return {"success": True, "base64": result, "mime_type": mime, "source_bytes": len(raw)}
            else:
                auto_path = make_output_path(input_path, "_b64", ".txt")
                Path(auto_path).write_text(result, encoding="ascii")
                return {"success": True, "output_path": auto_path, "note": "File >500 KB; saved to disk instead of inline.", "mime_type": mime}
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def base64_to_image(
        base64_string: Optional[str] = None,
        input_path: Optional[str] = None,
        output_path: str = "",
    ) -> Dict:
        """
        Decode a Base64 string (or a .txt file containing one) back to an image file.

        base64_string: raw base64 or data URI string.
        input_path: path to a .txt file containing the base64 string.
        output_path: destination image file path (required).
        """
        import base64, re
        try:
            if not output_path:
                return {"success": False, "error": "output_path is required."}
            if not base64_string and not input_path:
                return {"success": False, "error": "Provide either base64_string or input_path."}

            if input_path:
                src = resolve_input(input_path)
                b64_raw = src.read_text(encoding="ascii").strip()
            else:
                b64_raw = base64_string.strip()

            # Strip data URI prefix if present
            b64_raw = re.sub(r"^data:[^;]+;base64,", "", b64_raw)
            raw = base64.b64decode(b64_raw)

            dst = resolve_output(output_path)
            dst.write_bytes(raw)
            return {"success": True, "output_path": str(dst), "file_size_bytes": len(raw)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── HEIC / HEIF ─────────────────────────────────────────────────────────
    @app.tool()
    def heic_to_image(
        input_path: str,
        output_path: Optional[str] = None,
        output_format: str = "jpg",
        quality: int = 90,
    ) -> Dict:
        """
        Convert HEIC/HEIF images (Apple format) to JPG, PNG, or WebP.

        Requires pillow-heif (installed automatically).
        """
        try:
            from pillow_heif import register_heif_opener
            from PIL import Image
            register_heif_opener()

            src = resolve_input(input_path)
            fmt_map = {"jpg": ("JPEG", ".jpg"), "jpeg": ("JPEG", ".jpg"), "png": ("PNG", ".png"), "webp": ("WEBP", ".webp")}
            pil_fmt, ext = fmt_map.get(output_format.lower(), ("JPEG", ".jpg"))
            if not output_path:
                output_path = make_output_path(input_path, "", ext)
            dst = resolve_output(output_path)

            img = Image.open(src)
            if pil_fmt == "JPEG" and img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            kwargs = {"quality": quality} if pil_fmt in ("JPEG", "WEBP") else {}
            img.save(str(dst), format=pil_fmt, **kwargs)
            return {"success": True, "output_path": str(dst), "original_size": list(img.size), "output_format": pil_fmt}
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}
