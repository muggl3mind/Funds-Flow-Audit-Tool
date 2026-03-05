"""
Stage 6a: Serialize the full index to JSON.
Schema is backwards-compatible with the simple demo's output/index.json.
"""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Optional

from agent.config import DealConfig
from agent.exceptions.exception_classifier import DealException
from agent.matcher.llm_matcher import MatchResult
from agent.normalizers.funds_flow_normalizer import TabScopeResult


def write_index(
    results: list[MatchResult],
    exceptions: list[DealException],
    scope: TabScopeResult,
    config: DealConfig,
    output_path: Path,
) -> dict:

    line_items = [_serialize_result(r) for r in results]

    # Summary stats
    statuses = [r.status for r in results]
    matched_results = [r for r in results if r.status == "matched"]
    missing_results  = [r for r in results if r.status == "missing"]
    partial_results  = [r for r in results if r.status == "partial"]
    exception_results = [r for r in results if r.status == "exception"]

    total_supported = sum(
        r.line_item.total_amount or 0
        for r in results
        if r.status in ("matched", "partial") and r.matched_document
    )
    total_unsupported = sum(
        r.line_item.total_amount or 0
        for r in results
        if r.status in ("missing", "exception") and r.matched_document is None
    )

    amount_mismatches = sum(
        1 for r in results
        if r.amount_agrees is False
    )

    conf_scores = [r.confidence_score for r in results if r.confidence_score > 0]
    avg_conf = round(sum(conf_scores) / len(conf_scores), 4) if conf_scores else None

    doc = {
        "deal": config.deal_name,
        "closing_date": config.closing_date,
        "client_role": config.client_role,
        "funds_flow_file": config.funds_flow_path.name,
        "agent_version": config.agent_version,
        "indexed_at": str(date.today()),

        "fund_allocations": config.fund_allocations,

        "tabs_in_scope":  scope.in_scope,
        "tabs_skipped":   scope.skipped,
        "tabs_for_review": scope.for_review,
        "tab_reasoning":  scope.reasoning,

        "line_items": line_items,

        "exceptions": [_serialize_exception(e) for e in exceptions],

        "summary": {
            "indexed_at":              str(date.today()),
            "total_line_items":        len(results),
            "matched":                 len(matched_results),
            "partial":                 len(partial_results),
            "missing":                 len(missing_results),
            "exceptions":              len(exception_results),
            "amount_mismatches":       amount_mismatches,
            "total_supported_amount":  round(total_supported, 2),
            "total_unsupported_amount": round(total_unsupported, 2),
            "match_confidence_avg":    avg_conf,
        },
    }

    output_path.write_text(json.dumps(doc, indent=2))
    return doc


def _serialize_result(r: MatchResult) -> dict:
    item = r.line_item
    doc  = r.matched_document
    return {
        "status":              r.status,
        "source_tab":          item.source_tab,
        "source_row":          item.source_row,
        "funds_flow_ref":      item.ref_number or None,
        "description":         item.description,
        "vendor_hint":         item.vendor_hint or None,
        "funds_flow_amount":   item.total_amount,
        "fund_allocations":    item.fund_allocations,

        "document_file":       doc.file_name if doc else None,
        "document_vendor":     doc.vendor_name if doc else None,
        "document_invoice_number": doc.invoice_number if doc else None,
        "document_date":       doc.invoice_date if doc else None,
        "document_amount":     doc.total_amount if doc else None,
        "document_type":       doc.document_type if doc else None,

        "amount_agrees":       r.amount_agrees,
        "amount_discrepancy":  r.amount_discrepancy,
        "match_strength":      r.match_strength,
        "confidence_score":    r.confidence_score,
        "exception_flags":     r.exception_flags,
        "llm_reasoning":       r.llm_reasoning,
        "notes":               r.notes,
    }


def _serialize_exception(e: DealException) -> dict:
    return {
        "type":             e.exception_type,
        "severity":         e.severity,
        "source_ref":       e.source_ref,
        "description":      e.description,
        "suggested_action": e.suggested_action,
        "related_amount":   e.related_amount,
    }
