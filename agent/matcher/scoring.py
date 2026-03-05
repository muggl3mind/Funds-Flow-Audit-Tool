"""
Deterministic confidence scoring formula.
No LLM calls — purely math and string comparison.
"""
from __future__ import annotations

import re
from typing import Optional

from agent.normalizers.document_normalizer import DocumentRecord
from agent.normalizers.funds_flow_normalizer import FundsFlowLineItem

DOC_TYPE_SCORES = {
    "invoice": 1.0,
    "receipt": 0.9,
    "email":   0.7,
    "wire_confirmation": 0.8,
    "other":   0.5,
}

AMOUNT_TOLERANCE_EXACT  = 1.00        # within $1 → exact
AMOUNT_TOLERANCE_CLOSE  = 0.05        # within 5%
AMOUNT_TOLERANCE_NEAR   = 0.15        # within 15%


def compute_confidence(
    llm_score: float,
    llm_strength: str,
    line_item: FundsFlowLineItem,
    doc: Optional[DocumentRecord],
) -> float:
    if doc is None:
        return 0.0

    amount_bonus = _amount_score(line_item.total_amount, doc.total_amount)
    vendor_sim   = _vendor_similarity(line_item.vendor_hint, doc.vendor_name)
    doc_type_fit = DOC_TYPE_SCORES.get(doc.document_type, 0.5)

    score = (
        llm_score    * 0.50
        + amount_bonus * 0.25
        + vendor_sim   * 0.15
        + doc_type_fit * 0.10
    )
    return round(min(score, 1.0), 4)


def _amount_score(ff_amount: Optional[float], doc_amount: Optional[float]) -> float:
    if ff_amount is None or doc_amount is None:
        return 0.0
    if ff_amount == 0:
        return 1.0 if doc_amount == 0 else 0.0
    diff = abs(ff_amount - doc_amount)
    pct = diff / abs(ff_amount)
    if diff <= AMOUNT_TOLERANCE_EXACT:
        return 1.0
    if pct <= AMOUNT_TOLERANCE_CLOSE:
        return 0.80
    if pct <= AMOUNT_TOLERANCE_NEAR:
        return 0.50
    return 0.0


def _vendor_similarity(hint: str, vendor: Optional[str]) -> float:
    if not hint or not vendor:
        return 0.3     # neutral — no vendor info to penalise on
    a_tokens = _tokenize(hint)
    b_tokens = _tokenize(vendor)
    if not a_tokens or not b_tokens:
        return 0.3
    intersection = a_tokens & b_tokens
    union = a_tokens | b_tokens
    return len(intersection) / len(union)


def _tokenize(s: str) -> set[str]:
    STOP = {"llp", "lp", "inc", "corp", "co", "ltd", "the", "and", "&", "of"}
    tokens = set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    return tokens - STOP


def amount_discrepancy(ff_amount: Optional[float], doc_amount: Optional[float]) -> Optional[float]:
    if ff_amount is None or doc_amount is None:
        return None
    return round(doc_amount - ff_amount, 2)


def amounts_agree(ff_amount: Optional[float], doc_amount: Optional[float]) -> Optional[bool]:
    if ff_amount is None or doc_amount is None:
        return None
    return abs(ff_amount - doc_amount) <= AMOUNT_TOLERANCE_EXACT


def classify_amount_exception(ff_amount: Optional[float], doc_amount: Optional[float]) -> Optional[str]:
    """Return an exception code if the amounts diverge, else None."""
    if ff_amount is None or doc_amount is None:
        return None
    diff = doc_amount - ff_amount
    pct  = diff / ff_amount if ff_amount != 0 else 0
    if abs(diff) <= AMOUNT_TOLERANCE_EXACT:
        return None
    if diff < 0 and 0.10 <= abs(pct) <= 0.90:
        return "PARTIAL_INVOICE"
    if diff > 0:
        return "AMOUNT_OVER"
    return "AMOUNT_MISMATCH"
