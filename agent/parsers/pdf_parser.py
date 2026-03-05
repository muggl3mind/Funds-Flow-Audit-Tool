"""
Stage 3a: Extract raw text from PDF documents.
Tries PyMuPDF (fitz) first; falls back to PyPDF2.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass
class RawDocument:
    file_path: Path
    file_name: str
    extension: str
    text_content: str
    extraction_method: str   # "pymupdf" | "pypdf2" | "ocr_required" | "email"
    page_count: int


def parse_pdf(path: Path) -> RawDocument:
    path = Path(path)
    text, method, pages = _try_pymupdf(path)
    if not text.strip():
        text, method, pages = _try_pypdf2(path)
    if not text.strip():
        method = "ocr_required"

    return RawDocument(
        file_path=path,
        file_name=path.name,
        extension=path.suffix.lower(),
        text_content=text,
        extraction_method=method,
        page_count=pages,
    )


def _try_pymupdf(path: Path):
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(str(path))
        pages = doc.page_count
        parts = []
        for i, page in enumerate(doc):
            t = page.get_text("text")
            if t.strip():
                parts.append(f"--- PAGE {i + 1} ---\n{t}")
        doc.close()
        return "\n".join(parts), "pymupdf", pages
    except Exception:
        return "", "pymupdf_failed", 0


def _try_pypdf2(path: Path):
    try:
        import PyPDF2
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            pages = len(reader.pages)
            parts = []
            for i, page in enumerate(reader.pages):
                t = page.extract_text() or ""
                if t.strip():
                    parts.append(f"--- PAGE {i + 1} ---\n{t}")
        return "\n".join(parts), "pypdf2", pages
    except Exception:
        return "", "pypdf2_failed", 0
