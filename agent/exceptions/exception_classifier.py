"""
Stage 5: Post-matching exception classification.

Reviews the full result set and produces a structured exception list covering:
  - Missing documents
  - Amount mismatches / partial invoices
  - Low-confidence matches
  - Duplicate document assignments
  - Orphan documents (in folder, not matched to any line item)
  - Tabs flagged for review
  - OCR-required documents
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from agent.matcher.llm_matcher import MatchResult
from agent.normalizers.document_normalizer import DocumentRecord
from agent.normalizers.funds_flow_normalizer import TabScopeResult
from agent.utils.amount_utils import format_usd

SEVERITY = {
    "MISSING_DOCUMENT":        "HIGH",
    "AMOUNT_MISMATCH":         "HIGH",
    "AMOUNT_OVER":             "HIGH",
    "PARTIAL_INVOICE":         "MEDIUM",
    "DUPLICATE_DOCUMENT":      "HIGH",
    "LOW_CONFIDENCE_MATCH":    "MEDIUM",
    "UNMATCHED_DOCUMENT":      "LOW",
    "SUPERSEDED_INVOICE":      "INFO",
    "CUMULATIVE_BILLING":      "INFO",
    "SELLER_TAB_SKIPPED":      "INFO",
    "REVIEW_TAB":              "MEDIUM",
    "OCR_REQUIRED":            "MEDIUM",
    "SAME_VENDOR_MULTI_LINE":  "LOW",
}

ACTIONS = {
    "MISSING_DOCUMENT":     "Obtain the support document and add to the deal folder.",
    "AMOUNT_MISMATCH":      "Reconcile the difference between the funds flow amount and document amount.",
    "AMOUNT_OVER":          "Document exceeds funds flow amount — verify correct invoice version.",
    "PARTIAL_INVOICE":      "Confirm whether this is a retainer/progress billing; obtain final invoice.",
    "DUPLICATE_DOCUMENT":   "Same document matched to multiple line items — verify separate invoices exist.",
    "LOW_CONFIDENCE_MATCH": "Match confidence is below threshold — manually confirm this pairing.",
    "UNMATCHED_DOCUMENT":   "Document in folder was not matched to any line item — verify if relevant.",
    "SUPERSEDED_INVOICE":   "Earlier invoice superseded by cumulative statement — retain for records; no action required.",
    "CUMULATIVE_BILLING":   "Invoice is a cumulative statement including prior unpaid amounts — confirm total agrees with funds flow.",
    "SELLER_TAB_SKIPPED":   "Seller expense tab was excluded because client is the buyer — confirm this is correct.",
    "REVIEW_TAB":           "Tab classification was uncertain — human review required before finalizing.",
    "OCR_REQUIRED":         "Document is a scanned image PDF — extract text via OCR before matching.",
    "SAME_VENDOR_MULTI_LINE": "Same vendor appears multiple times — ensure each line item has its own invoice.",
}


@dataclass
class DealException:
    exception_type: str
    severity: str
    source_ref: str           # "Buyer Expenses:Row5" or document filename
    description: str          # human-readable summary
    suggested_action: str
    related_amount: Optional[float] = None


def classify_exceptions(
    results: list[MatchResult],
    all_documents: list[DocumentRecord],
    scope: TabScopeResult,
) -> list[DealException]:

    exceptions: list[DealException] = []

    matched_doc_names = {
        r.matched_document.file_name
        for r in results
        if r.matched_document
    }

    for result in results:
        item = result.line_item

        for flag in result.exception_flags:
            if flag == "MISSING_DOCUMENT":
                exceptions.append(DealException(
                    exception_type="MISSING_DOCUMENT",
                    severity=SEVERITY["MISSING_DOCUMENT"],
                    source_ref=item.source_row,
                    description=(
                        f"No supporting document found for '{item.description}' "
                        f"({format_usd(item.total_amount)})"
                    ),
                    suggested_action=ACTIONS["MISSING_DOCUMENT"],
                    related_amount=item.total_amount,
                ))

            elif flag in ("AMOUNT_MISMATCH", "AMOUNT_OVER", "PARTIAL_INVOICE"):
                doc = result.matched_document
                exceptions.append(DealException(
                    exception_type=flag,
                    severity=SEVERITY[flag],
                    source_ref=item.source_row,
                    description=(
                        f"'{item.description}': funds flow {format_usd(item.total_amount)} "
                        f"vs document {format_usd(doc.total_amount if doc else None)} "
                        f"(delta {format_usd(result.amount_discrepancy)})"
                    ),
                    suggested_action=ACTIONS[flag],
                    related_amount=result.amount_discrepancy,
                ))

            elif flag == "DUPLICATE_DOCUMENT":
                doc = result.matched_document
                exceptions.append(DealException(
                    exception_type="DUPLICATE_DOCUMENT",
                    severity=SEVERITY["DUPLICATE_DOCUMENT"],
                    source_ref=item.source_row,
                    description=(
                        f"'{item.description}': document '{doc.file_name if doc else '?'}' "
                        f"is assigned to multiple line items"
                    ),
                    suggested_action=ACTIONS["DUPLICATE_DOCUMENT"],
                ))

            elif flag == "LOW_CONFIDENCE_MATCH":
                doc = result.matched_document
                exceptions.append(DealException(
                    exception_type="LOW_CONFIDENCE_MATCH",
                    severity=SEVERITY["LOW_CONFIDENCE_MATCH"],
                    source_ref=item.source_row,
                    description=(
                        f"'{item.description}' matched to '{doc.file_name if doc else '?'}' "
                        f"with confidence {result.confidence_score:.0%} — below threshold"
                    ),
                    suggested_action=ACTIONS["LOW_CONFIDENCE_MATCH"],
                ))

            elif flag == "OCR_REQUIRED":
                exceptions.append(DealException(
                    exception_type="OCR_REQUIRED",
                    severity=SEVERITY["OCR_REQUIRED"],
                    source_ref=item.source_row,
                    description=f"Matched document requires OCR to extract text",
                    suggested_action=ACTIONS["OCR_REQUIRED"],
                ))

    # Also surface CUMULATIVE_BILLING flag from matched results
    for result in results:
        if "CUMULATIVE_BILLING" in result.exception_flags:
            doc = result.matched_document
            exceptions.append(DealException(
                exception_type="CUMULATIVE_BILLING",
                severity=SEVERITY["CUMULATIVE_BILLING"],
                source_ref=result.line_item.source_row,
                description=(
                    f"'{result.line_item.description}': matched to cumulative invoice "
                    f"'{doc.file_name if doc else '?'}' ({format_usd(doc.total_amount if doc else None)}). "
                    f"Earlier invoice(s) from same vendor are superseded."
                ),
                suggested_action=ACTIONS["CUMULATIVE_BILLING"],
                related_amount=doc.total_amount if doc else None,
            ))

    # Superseded and orphan documents
    for doc in all_documents:
        if doc.file_name not in matched_doc_names:
            if doc.document_type == "superseded":
                # Superseded by a cumulative invoice — informational only
                exceptions.append(DealException(
                    exception_type="SUPERSEDED_INVOICE",
                    severity=SEVERITY["SUPERSEDED_INVOICE"],
                    source_ref=doc.file_name,
                    description=(
                        f"'{doc.file_name}' ({doc.vendor_name}, "
                        f"{format_usd(doc.total_amount)}) is superseded by a later "
                        f"cumulative invoice. Retained for records."
                    ),
                    suggested_action=ACTIONS["SUPERSEDED_INVOICE"],
                    related_amount=doc.total_amount,
                ))
            else:
                # Genuine orphan
                exceptions.append(DealException(
                    exception_type="UNMATCHED_DOCUMENT",
                    severity=SEVERITY["UNMATCHED_DOCUMENT"],
                    source_ref=doc.file_name,
                    description=(
                        f"Document '{doc.file_name}' ({doc.vendor_name}, "
                        f"{format_usd(doc.total_amount)}) was not matched to any line item"
                    ),
                    suggested_action=ACTIONS["UNMATCHED_DOCUMENT"],
                    related_amount=doc.total_amount,
                ))

    # OCR-required documents
    for doc in all_documents:
        if doc.ocr_required:
            exceptions.append(DealException(
                exception_type="OCR_REQUIRED",
                severity=SEVERITY["OCR_REQUIRED"],
                source_ref=doc.file_name,
                description=f"'{doc.file_name}' is a scanned PDF — text could not be extracted",
                suggested_action=ACTIONS["OCR_REQUIRED"],
            ))

    # Review tabs
    for tab in scope.for_review:
        exceptions.append(DealException(
            exception_type="REVIEW_TAB",
            severity=SEVERITY["REVIEW_TAB"],
            source_ref=f"Tab: {tab}",
            description=(
                f"Tab '{tab}' could not be auto-classified as buyer or seller expense — "
                f"it was excluded from this run"
            ),
            suggested_action=ACTIONS["REVIEW_TAB"],
        ))

    # Seller tabs (informational audit log)
    for tab in scope.skipped:
        if any(kw in tab.lower() for kw in ("seller", "vendor", "target")):
            exceptions.append(DealException(
                exception_type="SELLER_TAB_SKIPPED",
                severity=SEVERITY["SELLER_TAB_SKIPPED"],
                source_ref=f"Tab: {tab}",
                description=(
                    f"Tab '{tab}' was identified as a seller expense tab and excluded "
                    f"because the client is the buyer"
                ),
                suggested_action=ACTIONS["SELLER_TAB_SKIPPED"],
            ))

    # Sort: HIGH first, then MEDIUM, INFO last
    order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2, "INFO": 3}
    exceptions.sort(key=lambda e: order.get(e.severity, 9))

    return exceptions
