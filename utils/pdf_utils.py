"""Shared PyMuPDF helpers."""
import io
from pathlib import Path
from typing import Optional
import fitz


def open_pdf(path: str, password: Optional[str] = None) -> fitz.Document:
    doc = fitz.open(str(path))
    if doc.is_encrypted:
        if not password:
            raise ValueError("PDF is encrypted. Provide a password via the 'password' parameter.")
        if not doc.authenticate(password):
            raise ValueError("Incorrect password for encrypted PDF.")
    return doc


def save_pdf(doc: fitz.Document, output_path: str, **kwargs) -> None:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path), **kwargs)


def pixmap_to_pil(pix: fitz.Pixmap):
    from PIL import Image
    mode = "RGBA" if pix.alpha else "RGB"
    return Image.frombytes(mode, (pix.width, pix.height), pix.samples)


def render_page(doc: fitz.Document, page_num: int, dpi: int = 150) -> fitz.Pixmap:
    page = doc[page_num]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    return page.get_pixmap(matrix=mat, alpha=False)
