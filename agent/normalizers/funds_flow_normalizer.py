"""
Stage 2: Funds flow normalization.

2a — Tab scope detection (rule-based + LLM fallback)
2b — Line item extraction from each in-scope sheet (LLM)
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Literal, Optional

from agent.config import DealConfig
from agent.parsers.excel_parser import RawSheet, sheet_to_prompt_text
from agent.utils.amount_utils import parse_amount
from agent.utils.claude_client import ClaudeClient
from agent.utils.logging_utils import RunLogger

TabDecision = Literal["IN_SCOPE", "SKIP", "REVIEW"]

BUYER_KW   = {"buyer", "purchaser", "acquirer", "acquiror", "acquirer's"}
SELLER_KW  = {"seller", "vendor", "target", "target co", "sellers", "company expenses",
               "seller's", "seller cost", "seller expenses"}
SKIP_KW    = {"wire", "instructions", "cover", "index", "toc", "table of contents",
               "contents", "notes to", "assumptions"}


@dataclass
class FundsFlowLineItem:
    source_tab: str
    source_row: str                     # e.g. "Buyer Expenses:Row5"
    description: str
    vendor_hint: str                    # best guess at vendor name, may be ""
    total_amount: Optional[float]
    fund_allocations: dict[str, float]  # {"Fund I": 450000.0, "Fund II": 300000.0}
    ref_number: str                     # explicit ref if present, else ""
    raw_notes: str


@dataclass
class TabScopeResult:
    decisions: dict[str, TabDecision]          # tab_name → decision
    reasoning: dict[str, str]                  # tab_name → rationale
    in_scope: list[str]
    skipped: list[str]
    for_review: list[str]


# ─── Tab Scope Detection ──────────────────────────────────────────────────────

def detect_tab_scope(
    sheets: list[RawSheet],
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
) -> TabScopeResult:

    decisions: dict[str, TabDecision] = {}
    reasoning: dict[str, str] = {}
    ambiguous: list[str] = []

    for sheet in sheets:
        decision, reason = _rule_classify(sheet.name, config.client_role)
        if decision == "AMBIGUOUS":
            ambiguous.append(sheet.name)
        else:
            decisions[sheet.name] = decision
            reasoning[sheet.name] = reason
            logger.info(f"Tab '{sheet.name}' → {decision} (rule-based)", stage="2a",
                        detail=reason)

    # LLM pass for ambiguous tabs only
    if ambiguous:
        llm_result = _llm_classify_tabs(ambiguous, sheets, config, client, logger)
        decisions.update(llm_result["decisions"])
        reasoning.update(llm_result["reasoning"])

    return TabScopeResult(
        decisions=decisions,
        reasoning=reasoning,
        in_scope=[t for t, d in decisions.items() if d == "IN_SCOPE"],
        skipped=[t for t, d in decisions.items() if d == "SKIP"],
        for_review=[t for t, d in decisions.items() if d == "REVIEW"],
    )


def _rule_classify(tab_name: str, client_role: str):
    name_lower = tab_name.lower().strip()
    tokens = set(name_lower.replace("-", " ").replace("&", " ").split())

    if tokens & SKIP_KW:
        return "SKIP", f"Tab name contains skip keyword(s): {tokens & SKIP_KW}"
    if client_role == "buyer" and tokens & SELLER_KW:
        return "SKIP", f"Client is buyer; tab name contains seller keyword(s): {tokens & SELLER_KW}"
    if client_role == "seller" and tokens & BUYER_KW:
        return "SKIP", f"Client is seller; tab name contains buyer keyword(s): {tokens & BUYER_KW}"
    if tokens & BUYER_KW:
        return "IN_SCOPE", f"Tab name contains buyer keyword(s): {tokens & BUYER_KW}"
    if tokens & SELLER_KW:
        return "SKIP", f"Tab name contains seller keyword(s): {tokens & SELLER_KW}"
    return "AMBIGUOUS", "No keyword match; requires LLM classification"


def _llm_classify_tabs(
    ambiguous_names: list[str],
    sheets: list[RawSheet],
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
) -> dict:
    # Build preview text for each ambiguous tab (first 30 non-blank rows)
    previews = {}
    for sheet in sheets:
        if sheet.name in ambiguous_names:
            previews[sheet.name] = sheet_to_prompt_text(sheet, max_rows=30)

    tab_blocks = "\n\n".join(
        f"TAB NAME: {name!r}\nCONTENT PREVIEW:\n{previews[name]}"
        for name in ambiguous_names
    )

    prompt = f"""You are analyzing an Excel funds flow workbook for a PE acquisition deal.
Deal: {config.deal_name}
Client role: {config.client_role.upper()} (the client is {config.client_role}ing the company)

The following tabs could NOT be classified by keyword rules. Review their content and classify each one.

{tab_blocks}

For each tab, classify as exactly one of:
  IN_SCOPE  — contains transaction costs that the {config.client_role} is responsible for paying
  SKIP      — seller costs the buyer doesn't owe, summary/cover/reference tabs, wire instruction tabs
  REVIEW    — genuinely ambiguous; needs human review before including

Return ONLY a JSON object in this exact format:
{{
  "decisions": {{
    "Tab Name 1": "IN_SCOPE",
    "Tab Name 2": "SKIP"
  }},
  "reasoning": {{
    "Tab Name 1": "One sentence explanation",
    "Tab Name 2": "One sentence explanation"
  }}
}}"""

    result = client.call_json(prompt, stage="2a")
    for name in ambiguous_names:
        d = result.get("decisions", {}).get(name, "REVIEW")
        r = result.get("reasoning", {}).get(name, "LLM classification")
        logger.info(f"Tab '{name}' → {d} (LLM)", stage="2a", detail=r)
    return result


# ─── Line Item Extraction ─────────────────────────────────────────────────────

def extract_line_items(
    sheets: list[RawSheet],
    scope: TabScopeResult,
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
) -> list[FundsFlowLineItem]:
    all_items: list[FundsFlowLineItem] = []

    for sheet in sheets:
        if sheet.name not in scope.in_scope:
            continue
        logger.info(f"Extracting line items from tab '{sheet.name}'", stage="2b")
        items = _extract_from_sheet(sheet, config, client, logger)
        logger.info(f"  → {len(items)} line item(s) extracted", stage="2b")
        all_items.extend(items)

    return all_items


def _extract_from_sheet(
    sheet: RawSheet,
    config: DealConfig,
    client: ClaudeClient,
    logger: RunLogger,
) -> list[FundsFlowLineItem]:

    sheet_text = sheet_to_prompt_text(sheet, max_rows=120)

    fund_names = list(config.fund_allocations.keys()) if config.fund_allocations else []
    fund_hint = (
        f"The funds are: {', '.join(fund_names)}. "
        f"Each line item may have a column per fund plus a Total column."
        if fund_names else
        "Each line item may have columns for individual funds and a Total."
    )

    prompt = f"""You are extracting transaction cost line items from a PE funds flow Excel tab.
Deal: {config.deal_name}  |  Client role: {config.client_role.upper()}

{fund_hint}

Below is the raw content of the tab, row by row. Rows tagged [SUBTOTAL] are rollup totals — exclude them as standalone line items. Rows that are section headers (text only, no dollar amounts) should also be excluded.

Extract ONLY the individual billable line items — one entry per actual cost the client is paying.

{sheet_text}

For each line item, extract:
  - description: the service/cost description
  - vendor_hint: the vendor or firm name if identifiable from description or notes (else "")
  - total_amount: the Total dollar amount as a number (null if not shown)
  - fund_allocations: dict of fund name → amount, e.g. {{"Fund I": 450000, "Fund II": 300000}}
  - ref_number: any explicit invoice/support reference in the row (else "")
  - raw_notes: any notes or comments from the row
  - source_row: "SheetName:RowN" where N is the row number (use the Row numbers shown above)

If a line item description spans two rows (description on one row, amounts on next), merge them into one entry.

Return ONLY a JSON array:
[
  {{
    "description": "Legal Counsel - Law Firm LLP",
    "vendor_hint": "Law Firm LLP",
    "total_amount": 750000,
    "fund_allocations": {{"Fund I": 450000, "Fund II": 300000}},
    "ref_number": "",
    "raw_notes": "Deal counsel; billed at closing",
    "source_row": "{sheet.name}:Row5"
  }}
]"""

    raw_items = client.call_json(prompt, stage="2b")
    if not isinstance(raw_items, list):
        logger.warn(f"LLM returned non-list for sheet '{sheet.name}'", stage="2b")
        return []

    result: list[FundsFlowLineItem] = []
    for item in raw_items:
        total = item.get("total_amount")
        if total is not None:
            total = parse_amount(total)

        fund_alloc_raw = item.get("fund_allocations") or {}
        fund_alloc = {k: parse_amount(v) or 0.0 for k, v in fund_alloc_raw.items()}

        result.append(FundsFlowLineItem(
            source_tab=sheet.name,
            source_row=item.get("source_row", f"{sheet.name}:unknown"),
            description=item.get("description", ""),
            vendor_hint=item.get("vendor_hint", ""),
            total_amount=total,
            fund_allocations=fund_alloc,
            ref_number=item.get("ref_number", ""),
            raw_notes=item.get("raw_notes", ""),
        ))

    return result
