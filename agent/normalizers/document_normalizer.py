"""
Stage 3b: Extract structured billing information from each support document.
Runs in parallel (ThreadPoolExecutor) and caches results to disk.
"""
from __future__ import annotations

import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from agent.config import DealConfig
from agent.parsers.email_parser import parse_document
from agent.parsers.pdf_parser import RawDocument
from agent.utils.amount_utils import parse_amount
from agent.utils.claude_client import ClaudeClient
from agent.utils.logging_utils import RunLogger

SUPPORTED_EXTENSIONS = {".pdf", ".eml", ".msg"}
MAX_TEXT_CHARS = 6000   # truncate to stay within prompt budget


@dataclass
class LineItemDetail:
    description: str
    amount: Optional[float]


@dataclass
class DocumentRecord:
    file_path: Path
    file_name: str
    vendor_name: Optional[str]
    invoice_number: Optional[str]
    invoice_date: Optional[str]
    total_amount: Optional[float]
    line_items: list[LineItemDetail]
    document_type: str                 # "invoice" | "email" | "receipt" | "wire_confirmation" | "other"
    deal_reference: Optional[str]
    extraction_method: str
    notes: str
    ocr_required: bool = False


def load_all_documents(
    documents_dir: Path,
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
    cache_path: Optional[Path] = None,
) -> list[DocumentRecord]:
    """Parse and extract all documents; use cache if available."""

    # Load cache if it exists
    cached: dict[str, dict] = {}
    if cache_path and cache_path.exists():
        try:
            cached = json.loads(cache_path.read_text())
            logger.info(f"Loaded document cache ({len(cached)} entries)", stage="3b")
        except Exception:
            cached = {}

    doc_files = [
        p for p in sorted(documents_dir.iterdir())
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    logger.info(f"Found {len(doc_files)} document(s) in {documents_dir}", stage="3a")

    # Determine which files need processing
    to_process: list[Path] = []
    records: list[DocumentRecord] = []

    for p in doc_files:
        if p.name in cached:
            records.append(_dict_to_record(cached[p.name], p))
            logger.info(f"  Cache hit: {p.name}", stage="3b")
        else:
            to_process.append(p)

    # Process uncached files in parallel
    if to_process:
        logger.info(f"Extracting {len(to_process)} document(s) via LLM (parallel)", stage="3b")
        new_records = _parallel_extract(to_process, config, client, logger)
        records.extend(new_records)

        # Update cache
        if cache_path:
            for rec in new_records:
                cached[rec.file_name] = _record_to_dict(rec)
            cache_path.write_text(json.dumps(cached, indent=2))

    return records


def _parallel_extract(
    paths: list[Path],
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
    max_workers: int = 4,
) -> list[DocumentRecord]:
    results: list[DocumentRecord] = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(_extract_one, p, config, client, logger): p
            for p in paths
        }
        for future in as_completed(futures):
            path = futures[future]
            try:
                rec = future.result()
                results.append(rec)
                logger.info(f"  Extracted: {path.name} → {rec.vendor_name} / "
                            f"{rec.total_amount}", stage="3b")
            except Exception as e:
                logger.error(f"  Failed: {path.name} — {e}", stage="3b")

    return results


def _extract_one(
    path: Path,
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
) -> DocumentRecord:
    raw: RawDocument = parse_document(path)

    if raw.extraction_method == "ocr_required":
        return DocumentRecord(
            file_path=path, file_name=path.name,
            vendor_name=None, invoice_number=None, invoice_date=None,
            total_amount=None, line_items=[], document_type="other",
            deal_reference=None, extraction_method="ocr_required",
            notes="Scanned PDF — no embedded text. OCR required.",
            ocr_required=True,
        )

    text = raw.text_content[:MAX_TEXT_CHARS]

    prompt = f"""You are extracting billing information from a support document for the PE deal "{config.deal_name}".

Document filename: {path.name}

Document text:
---
{text}
---

Extract the following fields. Use null for any field you cannot determine.

Return ONLY valid JSON:
{{
  "vendor_name": "exact vendor or firm name as it appears",
  "invoice_number": "invoice or reference number, or null",
  "invoice_date": "YYYY-MM-DD or closest approximation, or null",
  "total_amount": 123456.78,
  "line_items": [
    {{"description": "service description", "amount": 12345.67}}
  ],
  "document_type": "invoice" or "email" or "receipt" or "wire_confirmation" or "other",
  "deal_reference": "any deal or project name mentioned, or null",
  "notes": "any caveats, observations, or unusual aspects"
}}"""

    data = client.call_json(prompt, stage="3b")

    line_items = [
        LineItemDetail(
            description=li.get("description", ""),
            amount=parse_amount(li.get("amount")),
        )
        for li in (data.get("line_items") or [])
    ]

    return DocumentRecord(
        file_path=path,
        file_name=path.name,
        vendor_name=data.get("vendor_name"),
        invoice_number=data.get("invoice_number"),
        invoice_date=data.get("invoice_date"),
        total_amount=parse_amount(data.get("total_amount")),
        line_items=line_items,
        document_type=data.get("document_type", "other"),
        deal_reference=data.get("deal_reference"),
        extraction_method=raw.extraction_method,
        notes=data.get("notes", ""),
        ocr_required=False,
    )


def _record_to_dict(rec: DocumentRecord) -> dict:
    return {
        "file_name": rec.file_name,
        "vendor_name": rec.vendor_name,
        "invoice_number": rec.invoice_number,
        "invoice_date": rec.invoice_date,
        "total_amount": rec.total_amount,
        "line_items": [{"description": li.description, "amount": li.amount}
                       for li in rec.line_items],
        "document_type": rec.document_type,
        "deal_reference": rec.deal_reference,
        "extraction_method": rec.extraction_method,
        "notes": rec.notes,
        "ocr_required": rec.ocr_required,
    }


def _dict_to_record(d: dict, path: Path) -> DocumentRecord:
    return DocumentRecord(
        file_path=path,
        file_name=d["file_name"],
        vendor_name=d.get("vendor_name"),
        invoice_number=d.get("invoice_number"),
        invoice_date=d.get("invoice_date"),
        total_amount=d.get("total_amount"),
        line_items=[
            LineItemDetail(li["description"], li.get("amount"))
            for li in (d.get("line_items") or [])
        ],
        document_type=d.get("document_type", "other"),
        deal_reference=d.get("deal_reference"),
        extraction_method=d.get("extraction_method", "cached"),
        notes=d.get("notes", ""),
        ocr_required=d.get("ocr_required", False),
    )
