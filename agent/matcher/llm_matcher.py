"""
Stage 4: LLM-driven matching of funds flow line items to support documents.

For each line item:
  1. Pre-filter candidate documents (amount proximity + vendor token overlap)
  2. LLM call: reason over candidates and pick the best match
  3. Score: deterministic confidence formula
  4. Detect duplicate assignments after all items are matched
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Optional

from collections import defaultdict

from agent.config import DealConfig
from agent.matcher.scoring import (
    amount_discrepancy, amounts_agree, classify_amount_exception,
    compute_confidence, _vendor_similarity, _amount_score,
)
from agent.normalizers.document_normalizer import DocumentRecord
from agent.normalizers.funds_flow_normalizer import FundsFlowLineItem
from agent.utils.amount_utils import format_usd
from agent.utils.claude_client import ClaudeClient
from agent.utils.logging_utils import RunLogger

AMOUNT_WINDOW = 0.20   # pre-filter: doc within 20% of ff amount (or no ff amount)
VENDOR_MIN    = 0.15   # pre-filter: Jaccard token overlap threshold


def build_cumulative_map(documents: list[DocumentRecord]) -> dict[str, list[DocumentRecord]]:
    """
    Group documents by normalised vendor name where multiple invoices exist from the
    same vendor. Used to detect cumulative / superseding billing patterns.

    Returns: { normalised_vendor_key: [doc_a, doc_b, ...] }
    Only includes vendors with 2+ documents.
    """
    by_vendor: dict[str, list[DocumentRecord]] = defaultdict(list)
    for doc in documents:
        if doc.vendor_name:
            key = _normalise_vendor(doc.vendor_name)
            by_vendor[key].append(doc)
    return {k: v for k, v in by_vendor.items() if len(v) > 1}


def _normalise_vendor(name: str) -> str:
    """Lowercase + strip punctuation for grouping purposes."""
    import re
    return re.sub(r"[^a-z0-9 ]", " ", name.lower()).split()[0:3].__str__()


@dataclass
class MatchResult:
    line_item: FundsFlowLineItem
    matched_document: Optional[DocumentRecord]
    match_strength: str          # STRONG_MATCH | PARTIAL_MATCH | WEAK_MATCH | NO_MATCH | UNMATCHED
    confidence_score: float
    status: str                  # "matched" | "missing" | "partial" | "exception"
    amount_agrees: Optional[bool]
    amount_discrepancy: Optional[float]
    exception_flags: list[str]
    llm_reasoning: str
    notes: str


def match_all(
    line_items: list[FundsFlowLineItem],
    documents: list[DocumentRecord],
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
    max_workers: int = 4,
) -> list[MatchResult]:
    logger.info(f"Matching {len(line_items)} line item(s) against "
                f"{len(documents)} document(s)", stage="4")

    # Build cumulative billing map once — shared across all parallel calls
    cumulative_map = build_cumulative_map(documents)
    if cumulative_map:
        logger.info(
            f"  Cumulative billing groups detected: "
            f"{[k for k in cumulative_map]}",
            stage="4",
        )

    # Parallel matching
    results: list[MatchResult] = [None] * len(line_items)  # type: ignore

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                _match_one, item, documents, cumulative_map, config, client, logger
            ): idx
            for idx, item in enumerate(line_items)
        }
        for future in as_completed(futures):
            idx = futures[future]
            try:
                results[idx] = future.result()
            except Exception as e:
                item = line_items[idx]
                logger.error(f"Match failed for '{item.description}': {e}", stage="4")
                results[idx] = _unmatched(item, str(e))

    results = [r for r in results if r is not None]

    # Detect duplicate document assignments
    _flag_duplicates(results, logger)

    # Mark superseded invoices (earlier invoices baked into a later cumulative one)
    _flag_superseded(results, cumulative_map, logger)

    return results


def _match_one(
    item: FundsFlowLineItem,
    documents: list[DocumentRecord],
    cumulative_map: dict[str, list[DocumentRecord]],
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
) -> MatchResult:

    candidates = _prefilter(item, documents)
    if not candidates:
        logger.info(f"  No candidates for '{item.description}' — marking missing", stage="4")
        return _unmatched(item, "No candidate documents passed pre-filter")

    # Detect whether any vendor group has multiple invoices among the candidates
    cumulative_note = _build_cumulative_note(candidates, cumulative_map, item)

    # Build LLM prompt
    cand_block = _format_candidates(candidates)
    prompt = f"""You are a PE deal accountant verifying transaction cost support documents.

FUNDS FLOW LINE ITEM:
  Description:  {item.description}
  Vendor hint:  {item.vendor_hint or '(not specified)'}
  Total amount: {format_usd(item.total_amount)}
  Fund split:   {_fmt_alloc(item.fund_allocations)}
  Source tab:   {item.source_tab}
  Notes:        {item.raw_notes or '(none)'}

CANDIDATE DOCUMENTS ({len(candidates)} found):
{cand_block}
{cumulative_note}
For EACH candidate, rate the match:
  STRONG_MATCH  — vendor, service description, AND amount all clearly align
  PARTIAL_MATCH — vendor and description align but amount differs (e.g. retainer, split billing)
  WEAK_MATCH    — some overlap but significant uncertainty
  NO_MATCH      — clearly wrong vendor or unrelated service

Then identify the single best match (or null if none are good enough).

Return ONLY valid JSON:
{{
  "candidate_ratings": [
    {{"index": 1, "strength": "STRONG_MATCH", "reasoning": "..."}}
  ],
  "best_match_index": 1,
  "match_strength": "STRONG_MATCH",
  "confidence_score": 0.97,
  "reasoning": "One paragraph explanation",
  "amount_agrees": true,
  "exception_flags": []
}}

Exception flag codes (include any that apply):
  PARTIAL_INVOICE        — doc amount is lower; may be a retainer or progress billing
  AMOUNT_OVER            — doc amount exceeds funds flow amount
  AMOUNT_MISMATCH        — amounts differ with no clear explanation
  CUMULATIVE_BILLING     — invoice is cumulative and supersedes an earlier invoice
  SAME_VENDOR_MULTI_LINE — same vendor appears multiple times in the funds flow
  LOW_CONFIDENCE_MATCH   — uncertain match; human review recommended
  NO_MATCH               — no suitable document found"""

    data = client.call_json(prompt, stage="4")

    best_idx = data.get("best_match_index")
    strength = data.get("match_strength", "NO_MATCH")
    llm_conf = float(data.get("confidence_score") or 0.0)
    reasoning = data.get("reasoning", "")
    llm_flags: list[str] = data.get("exception_flags") or []

    best_doc: Optional[DocumentRecord] = None
    if best_idx is not None:
        try:
            best_doc = candidates[int(best_idx) - 1]  # LLM uses 1-indexed
        except (IndexError, ValueError):
            best_doc = None

    conf = compute_confidence(llm_conf, strength, item, best_doc)
    status, flags = _derive_status(strength, conf, item, best_doc, llm_flags, config)

    disc = amount_discrepancy(item.total_amount, best_doc.total_amount if best_doc else None)
    agrees = amounts_agree(item.total_amount, best_doc.total_amount if best_doc else None)

    # Amount-based exception
    amt_exc = classify_amount_exception(
        item.total_amount, best_doc.total_amount if best_doc else None
    )
    if amt_exc and amt_exc not in flags:
        flags.append(amt_exc)

    logger.info(
        f"  '{item.description[:45]}' → "
        f"{'[' + best_doc.file_name + ']' if best_doc else '[no match]'} "
        f"| {strength} | conf={conf:.2f} | status={status}",
        stage="4",
    )

    notes = reasoning
    if disc and disc != 0:
        notes += f" | Amount delta: {format_usd(disc)}"

    return MatchResult(
        line_item=item,
        matched_document=best_doc,
        match_strength=strength,
        confidence_score=conf,
        status=status,
        amount_agrees=agrees,
        amount_discrepancy=disc,
        exception_flags=flags,
        llm_reasoning=reasoning,
        notes=notes,
    )


def _prefilter(item: FundsFlowLineItem, documents: list[DocumentRecord]) -> list[DocumentRecord]:
    """Return candidate documents that pass amount or vendor proximity filters."""
    if not documents:
        return []

    candidates = []
    for doc in documents:
        amount_ok = False
        vendor_ok = False

        if item.total_amount is not None and doc.total_amount is not None:
            pct = abs(item.total_amount - doc.total_amount) / max(abs(item.total_amount), 1)
            amount_ok = pct <= AMOUNT_WINDOW
        elif item.total_amount is None or doc.total_amount is None:
            amount_ok = True   # can't eliminate on amount alone if either is missing

        if item.vendor_hint and doc.vendor_name:
            vendor_ok = _vendor_similarity(item.vendor_hint, doc.vendor_name) >= VENDOR_MIN

        if amount_ok or vendor_ok:
            candidates.append(doc)

    # If pre-filter is too aggressive (e.g. small deal), fall back to all docs
    if not candidates:
        candidates = documents[:10]   # cap at 10 to keep prompt manageable

    return candidates


def _format_candidates(docs: list[DocumentRecord]) -> str:
    lines = []
    for i, doc in enumerate(docs, start=1):
        lines.append(
            f"[{i}] {doc.file_name}\n"
            f"    Vendor:     {doc.vendor_name or '(unknown)'}\n"
            f"    Invoice #:  {doc.invoice_number or '(none)'}\n"
            f"    Date:       {doc.invoice_date or '(unknown)'}\n"
            f"    Amount:     {format_usd(doc.total_amount)}\n"
            f"    Type:       {doc.document_type}\n"
            f"    Deal ref:   {doc.deal_reference or '(none)'}\n"
            f"    Notes:      {doc.notes[:120] if doc.notes else '(none)'}"
        )
    return "\n\n".join(lines)


def _fmt_alloc(alloc: dict) -> str:
    if not alloc:
        return "(not specified)"
    return "  |  ".join(f"{k}: {format_usd(v)}" for k, v in alloc.items())


def _derive_status(
    strength: str,
    conf: float,
    item: FundsFlowLineItem,
    doc: Optional[DocumentRecord],
    llm_flags: list[str],
    config: DealConfig,
) -> tuple[str, list[str]]:
    flags = list(llm_flags)
    threshold = config.match_confidence_threshold

    if doc is None or strength == "NO_MATCH":
        return "missing", flags

    if strength == "STRONG_MATCH":
        if conf >= 0.85:
            return "matched", flags
        if conf >= threshold:
            flags.append("LOW_CONFIDENCE_MATCH")
            return "matched", flags
        flags.append("LOW_CONFIDENCE_MATCH")
        return "exception", flags

    if strength == "PARTIAL_MATCH":
        if "PARTIAL_INVOICE" not in flags:
            flags.append("PARTIAL_INVOICE")
        return "partial", flags

    if strength == "WEAK_MATCH":
        flags.append("LOW_CONFIDENCE_MATCH")
        return "exception", flags

    return "missing", flags


def _flag_duplicates(results: list[MatchResult], logger: RunLogger):
    """If the same document is matched to multiple line items, flag both."""
    seen: dict[str, list[int]] = {}
    for i, r in enumerate(results):
        if r.matched_document:
            key = r.matched_document.file_name
            seen.setdefault(key, []).append(i)

    for fname, indices in seen.items():
        if len(indices) > 1:
            for i in indices:
                if "DUPLICATE_DOCUMENT" not in results[i].exception_flags:
                    results[i].exception_flags.append("DUPLICATE_DOCUMENT")
                results[i].status = "exception"
            logger.warn(
                f"Document '{fname}' matched to {len(indices)} line items — flagged as DUPLICATE",
                stage="4",
            )


def _unmatched(item: FundsFlowLineItem, reason: str) -> MatchResult:
    return MatchResult(
        line_item=item,
        matched_document=None,
        match_strength="NO_MATCH",
        confidence_score=0.0,
        status="missing",
        amount_agrees=None,
        amount_discrepancy=None,
        exception_flags=["MISSING_DOCUMENT"],
        llm_reasoning=reason,
        notes=reason,
    )


def _build_cumulative_note(
    candidates: list[DocumentRecord],
    cumulative_map: dict[str, list[DocumentRecord]],
    item: FundsFlowLineItem,
) -> str:
    """
    If the candidate list contains multiple invoices from the same vendor,
    inject an explicit note to the LLM prompt explaining cumulative billing.

    This prevents the LLM from incorrectly flagging the cumulative invoice as
    AMOUNT_OVER or matching to the superseded earlier invoice instead.
    """
    candidate_names = {d.file_name for d in candidates}
    multi_vendor_groups = []

    for vendor_key, group in cumulative_map.items():
        group_in_candidates = [d for d in group if d.file_name in candidate_names]
        if len(group_in_candidates) >= 2:
            # Sort by amount ascending — smaller amount is likely the earlier invoice
            sorted_group = sorted(
                group_in_candidates,
                key=lambda d: d.total_amount or 0
            )
            amounts = [format_usd(d.total_amount) for d in sorted_group]
            names   = [d.file_name for d in sorted_group]
            multi_vendor_groups.append((sorted_group[0].vendor_name, names, amounts))

    if not multi_vendor_groups:
        return ""

    lines = ["\n⚠ CUMULATIVE BILLING ALERT — IMPORTANT MATCHING INSTRUCTIONS:"]
    for vendor, names, amounts in multi_vendor_groups:
        lines.append(
            f"  Vendor '{vendor}' has {len(names)} invoices in the candidate list:\n"
            + "\n".join(f"    • {n}  ({a})" for n, a in zip(names, amounts))
        )
        lines.append(
            "  Law firms and advisors commonly issue CUMULATIVE STATEMENTS where each\n"
            "  new invoice includes all prior unpaid amounts. The most recent invoice\n"
            "  with the highest amount is typically the CONTROLLING document (it\n"
            "  supersedes earlier invoices). The earlier smaller invoice(s) are\n"
            "  SUPERSEDED — do not match the funds flow to a superseded invoice.\n"
            f"  If the funds flow amount agrees with the LARGEST invoice ({amounts[-1]}),\n"
            "  match to that invoice and flag it as CUMULATIVE_BILLING.\n"
            "  DO NOT flag a STRONG_MATCH as AMOUNT_OVER just because another smaller\n"
            "  invoice from the same vendor also exists.\n"
        )
    return "\n".join(lines) + "\n"


def _flag_superseded(
    results: list[MatchResult],
    cumulative_map: dict[str, list[DocumentRecord]],
    logger: RunLogger,
):
    """
    After matching, identify invoices that were NOT chosen as the best match
    for their vendor group. If the vendor group has a matched (winning) invoice,
    the unchosen smaller invoice is likely SUPERSEDED by the cumulative one.

    Adds a synthetic MatchResult with status="superseded" for each superseded doc,
    so it appears in the audit trail / unmatched docs sheet with the right label.
    """
    matched_docs = {r.matched_document.file_name for r in results if r.matched_document}

    for vendor_key, group in cumulative_map.items():
        # Is any member of this group the winning match for a line item?
        winner = next(
            (d for d in group if d.file_name in matched_docs),
            None,
        )
        if winner is None:
            continue  # no invoice from this vendor was matched — handled as MISSING

        for doc in group:
            if doc.file_name == winner.file_name:
                continue  # this is the winner, not superseded
            if doc.file_name in matched_docs:
                continue  # also matched to something (shouldn't happen, but safe)

            # This doc is superseded — tag it so exception_classifier picks it up
            doc.notes = (
                f"[SUPERSEDED] This invoice ({format_usd(doc.total_amount)}) is included in "
                f"the cumulative invoice '{winner.file_name}' ({format_usd(winner.total_amount)}). "
                f"The cumulative invoice is the controlling document. "
                + (doc.notes or "")
            )
            # Mark on the doc object so exception_classifier can detect it
            doc.document_type = "superseded"
            logger.info(
                f"  '{doc.file_name}' marked SUPERSEDED by '{winner.file_name}'",
                stage="4",
            )
