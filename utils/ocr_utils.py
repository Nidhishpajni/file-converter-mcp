"""Lazy EasyOCR loader and coordinate-mapping utilities."""
from typing import List, Tuple

_reader_cache: dict = {}

OCR_LANGUAGES = {
    "en": "English", "ch_sim": "Chinese (Simplified)", "ch_tra": "Chinese (Traditional)",
    "hi": "Hindi", "ar": "Arabic", "fr": "French", "de": "German", "es": "Spanish",
    "pt": "Portuguese", "it": "Italian", "ja": "Japanese", "ko": "Korean",
    "ru": "Russian", "nl": "Dutch", "pl": "Polish", "tr": "Turkish",
    "vi": "Vietnamese", "th": "Thai", "id": "Indonesian", "sv": "Swedish",
}


def get_ocr_reader(languages: List[str], gpu: bool = False):
    key = tuple(sorted(languages))
    if key not in _reader_cache:
        import easyocr
        _reader_cache[key] = easyocr.Reader(list(languages), gpu=gpu)
    return _reader_cache[key]


def bbox_to_points(bbox: List, dpi: int) -> Tuple[float, float, float, float]:
    scale = 72.0 / dpi
    xs = [p[0] for p in bbox]
    ys = [p[1] for p in bbox]
    x0, y0 = min(xs) * scale, min(ys) * scale
    x1, y1 = max(xs) * scale, max(ys) * scale
    return x0, y0, x1, y1
