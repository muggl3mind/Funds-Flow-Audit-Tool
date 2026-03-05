"""
Stage 6c: Rename and index support documents.

After matching is complete, copy every support document into
  <output_dir>/documents_indexed/

using the naming convention:
  FF01 - Vendor Name - INV-2026-001.pdf   (matched)
  UNMATCHED - OrphanDoc_Q2_2026.pdf       (in folder but not matched)

This runs automatically at the end of every agent run so that the
documents_indexed/ folder always mirrors the annotated Excel.
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Optional

from agent.matcher.llm_matcher import MatchResult
from agent.normalizers.document_normalizer import DocumentRecord
from agent.utils.logging_utils import RunLogger


# Characters that are illegal in filenames on macOS / Windows / Linux
_ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _safe(text: str, max_len: int = 40) -> str:
    """Sanitise a string for use as part of a filename."""
    s = _ILLEGAL.sub("-", text)
    s = re.sub(r"-{2,}", "-", s).strip("- ")
    return s[:max_len]


def _ff_label(n: int) -> str:
    return f"FF{n:02d}"


def rename_and_index(
    results: list[MatchResult],
    all_documents: list[DocumentRecord],
    output_dir: Path,
    logger: RunLogger,
) -> list[dict]:
    """
    Copy matched documents to <output_dir>/documents_indexed/ with FF-numbered names.
    Unmatched documents are copied into a UNMATCHED/ sub-folder.

    Returns a list of dicts describing what was renamed (written to run log).
    """
    index_dir     = output_dir / "documents_indexed"
    unmatched_dir = index_dir / "UNMATCHED"
    index_dir.mkdir(parents=True, exist_ok=True)
    unmatched_dir.mkdir(parents=True, exist_ok=True)

    log: list[dict] = []
    matched_files: set[str] = set()

    # ── Matched / partial / missing results ───────────────────────────────────
    counter = 1
    for result in results:
        doc = result.matched_document

        if doc is None:
            # Missing — no file to copy, just record it
            log.append({
                "ff_ref":    _ff_label(counter),
                "status":    "MISSING",
                "original":  None,
                "renamed":   None,
                "vendor":    result.line_item.vendor_hint or result.line_item.description,
                "inv_no":    None,
            })
            counter += 1
            continue

        matched_files.add(doc.file_name)

        # Build the new filename
        vendor_part = _safe(doc.vendor_name or result.line_item.vendor_hint or "Unknown", 36)
        inv_part    = _safe(doc.invoice_number or "no-ref", 24)
        suffix      = doc.file_path.suffix or ".pdf"
        new_name    = f"{_ff_label(counter)} - {vendor_part} - {inv_part}{suffix}"

        dst = index_dir / new_name
        _copy(doc.file_path, dst, logger)

        log.append({
            "ff_ref":   _ff_label(counter),
            "status":   result.status.upper(),
            "original": doc.file_name,
            "renamed":  new_name,
            "vendor":   doc.vendor_name,
            "inv_no":   doc.invoice_number,
        })
        counter += 1

    # ── Unmatched documents (in folder but not matched to any line item) ───────
    for doc in all_documents:
        if doc.file_name in matched_files:
            continue
        new_name = f"UNMATCHED - {doc.file_name}"
        dst = unmatched_dir / new_name
        _copy(doc.file_path, dst, logger)
        log.append({
            "ff_ref":   "—",
            "status":   "UNMATCHED",
            "original": doc.file_name,
            "renamed":  f"UNMATCHED/{new_name}",
            "vendor":   doc.vendor_name,
            "inv_no":   doc.invoice_number,
        })

    # ── Log summary ───────────────────────────────────────────────────────────
    n_ok        = sum(1 for e in log if e["status"] not in ("MISSING", "UNMATCHED"))
    n_missing   = sum(1 for e in log if e["status"] == "MISSING")
    n_unmatched = sum(1 for e in log if e["status"] == "UNMATCHED")

    logger.info(
        f"Document index: {n_ok} renamed, {n_missing} missing, "
        f"{n_unmatched} unmatched  →  {index_dir}",
        stage="6c",
    )
    for entry in log:
        icon = ("✓" if entry["status"] not in ("MISSING", "UNMATCHED")
                else ("✗" if entry["status"] == "MISSING" else "↩"))
        orig = entry["original"] or "(no document)"
        renamed = f"  →  {entry['renamed']}" if entry["renamed"] else ""
        logger.info(f"  {icon} {entry['ff_ref']:5}  {orig}{renamed}", stage="6c")

    return log


def _copy(src: Path, dst: Path, logger: RunLogger):
    try:
        shutil.copy2(src, dst)
    except Exception as e:
        logger.warning(f"  Could not copy {src.name}: {e}", stage="6c")
