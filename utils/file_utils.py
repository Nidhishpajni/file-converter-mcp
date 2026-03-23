"""Shared file path utilities."""
from pathlib import Path
from typing import Optional


def resolve_input(path: str) -> Path:
    p = Path(path).resolve()
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {p}")
    if not p.is_file():
        raise ValueError(f"Input path is not a file: {p}")
    return p


def resolve_output(path: str) -> Path:
    p = Path(path).resolve()
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def validate_input(path: str) -> Optional[str]:
    try:
        resolve_input(path)
        return None
    except (FileNotFoundError, ValueError) as e:
        return str(e)


def make_output_path(input_path: str, suffix: str, ext: Optional[str] = None) -> str:
    p = Path(input_path).resolve()
    stem = p.stem + suffix
    extension = ext if ext else p.suffix
    return str(p.parent / (stem + extension))
