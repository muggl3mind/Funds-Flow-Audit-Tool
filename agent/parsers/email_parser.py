"""
Stage 3a: Extract text from raw email files (.eml, .msg).
PDFs that happen to contain email content are handled by pdf_parser and
recognized by the document normalizer from their text structure.
"""
from __future__ import annotations

import email as email_lib
import html
import re
from pathlib import Path

from agent.parsers.pdf_parser import RawDocument


def parse_email_file(path: Path) -> RawDocument:
    path = Path(path)
    ext = path.suffix.lower()

    if ext == ".eml":
        text, method = _parse_eml(path)
    elif ext == ".msg":
        text, method = _parse_msg(path)
    else:
        text, method = "", "unsupported"

    return RawDocument(
        file_path=path,
        file_name=path.name,
        extension=ext,
        text_content=text,
        extraction_method=method,
        page_count=1,
    )


def _parse_eml(path: Path):
    try:
        with open(path, "rb") as f:
            msg = email_lib.message_from_bytes(f.read())

        headers = (
            f"FROM: {msg.get('From', '')}\n"
            f"TO: {msg.get('To', '')}\n"
            f"DATE: {msg.get('Date', '')}\n"
            f"SUBJECT: {msg.get('Subject', '')}\n"
            "---\n"
        )

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ct = part.get_content_type()
                if ct == "text/plain":
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                    break
                if ct == "text/html" and not body:
                    raw_html = part.get_payload(decode=True).decode("utf-8", errors="replace")
                    body = _strip_html(raw_html)
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode("utf-8", errors="replace")

        return headers + body, "eml"
    except Exception as e:
        return f"[EML parse error: {e}]", "eml_failed"


def _parse_msg(path: Path):
    try:
        import extract_msg
        msg = extract_msg.Message(str(path))
        headers = (
            f"FROM: {msg.sender}\n"
            f"TO: {msg.to}\n"
            f"DATE: {msg.date}\n"
            f"SUBJECT: {msg.subject}\n"
            "---\n"
        )
        body = msg.body or ""
        msg.close()
        return headers + body, "msg"
    except ImportError:
        return "[extract-msg not installed; cannot parse .msg file]", "msg_unavailable"
    except Exception as e:
        return f"[MSG parse error: {e}]", "msg_failed"


def _strip_html(raw: str) -> str:
    text = re.sub(r"<[^>]+>", " ", raw)
    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_document(path: Path) -> RawDocument:
    """
    Dispatcher: route a file to the right parser based on extension.
    Imported by document_normalizer.
    """
    ext = path.suffix.lower()
    if ext in (".eml", ".msg"):
        return parse_email_file(path)
    elif ext == ".pdf":
        from agent.parsers.pdf_parser import parse_pdf
        return parse_pdf(path)
    else:
        # Try PDF parser as fallback
        from agent.parsers.pdf_parser import parse_pdf
        return parse_pdf(path)
