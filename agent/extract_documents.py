"""
agent/extract_documents.py — Extract text from all PDFs in a documents folder.

Outputs a JSON array where each entry has the filename and full extracted text,
ready for Claude Code to read and reason over.

Usage:
  python agent/extract_documents.py <documents_dir> <output_json>
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


def _extract_pdf(path: Path) -> str:
    # Try PyMuPDF first (best quality)
    try:
        import fitz
        doc = fitz.open(str(path))
        text = "\n".join(page.get_text() for page in doc)
        doc.close()
        if text.strip():
            return text.strip()
    except Exception:
        pass

    # Fallback: pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        if text.strip():
            return text.strip()
    except Exception:
        pass

    return "[text extraction failed]"


def main():
    if len(sys.argv) < 3:
        print("Usage: python agent/extract_documents.py <documents_dir> <output_json>")
        sys.exit(1)

    docs_dir    = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    results = []
    for pdf in sorted(docs_dir.glob("*.pdf")):
        text = _extract_pdf(pdf)
        results.append({
            "filename":   pdf.name,
            "size_bytes": pdf.stat().st_size,
            "text":       text,
        })
        print(f"  {pdf.name}: {len(text)} chars extracted")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, indent=2))
    print(f"\n{len(results)} documents → {output_path}")


if __name__ == "__main__":
    main()
