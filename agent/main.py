"""
Funds Flow Indexer — Production Agent
======================================
Usage:
  python -m agent.main \\
    --deal "Project Titan V2" \\
    --closing-date 2026-03-03 \\
    --client-role buyer \\
    --funds-flow deals/project_titan_v2/funds_flow.xlsx \\
    --documents deals/project_titan_v2/documents/ \\
    --output deals/project_titan_v2/run_output/ \\
    [--fund "Fund I=0.60" --fund "Fund II=0.40"] \\
    [--model claude-opus-4-6] \\
    [--threshold 0.75]
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import date
from pathlib import Path

from agent.config import DealConfig
from agent.exceptions.exception_classifier import classify_exceptions
from agent.matcher.llm_matcher import match_all
from agent.normalizers.document_normalizer import load_all_documents
from agent.normalizers.funds_flow_normalizer import detect_tab_scope, extract_line_items
from agent.output.document_renamer import rename_and_index
from agent.output.excel_writer import annotate_client_file
from agent.output.json_writer import write_index
from agent.parsers.excel_parser import parse_workbook
from agent.utils.claude_client import ClaudeClient
from agent.utils.logging_utils import RunLogger


def main():
    args = _parse_args()
    run(
        deal_name=args.deal,
        closing_date=args.closing_date,
        client_role=args.client_role,
        funds_flow_path=Path(args.funds_flow),
        documents_dir=Path(args.documents),
        output_dir=Path(args.output),
        fund_allocations=_parse_allocations(args.fund),
        model=args.model,
        threshold=args.threshold,
    )


def run(
    deal_name: str,
    closing_date: str,
    client_role: str,
    funds_flow_path: Path,
    documents_dir: Path,
    output_dir: Path,
    fund_allocations: dict,
    model: str = "claude-opus-4-6",
    threshold: float = 0.75,
):
    t_start = time.time()
    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "run.log"
    logger = RunLogger(log_path=log_path, verbose=True)

    logger.info(f"Starting run: {deal_name}", stage="INIT")
    logger.info(f"Funds flow: {funds_flow_path}", stage="INIT")
    logger.info(f"Documents: {documents_dir}", stage="INIT")

    config = DealConfig(
        deal_name=deal_name,
        closing_date=closing_date,
        client_role=client_role,
        funds_flow_path=funds_flow_path,
        documents_dir=documents_dir,
        output_dir=output_dir,
        fund_allocations=fund_allocations,
        match_confidence_threshold=threshold,
        anthropic_model=model,
    )

    client = ClaudeClient(
        api_key=config.anthropic_api_key,
        model=config.anthropic_model,
        logger=logger,
    )

    # ── Stage 1: Parse Excel ──────────────────────────────────────────────────
    logger.info("Parsing funds flow Excel...", stage="1")
    sheets = parse_workbook(funds_flow_path)
    logger.info(f"  Found {len(sheets)} sheet(s): {[s.name for s in sheets]}", stage="1")

    # ── Stage 2a: Tab scope detection ────────────────────────────────────────
    logger.info("Detecting tab scope...", stage="2a")
    scope = detect_tab_scope(sheets, config, client, logger)
    logger.info(
        f"  In scope: {scope.in_scope} | Skipped: {scope.skipped} | Review: {scope.for_review}",
        stage="2a",
    )

    # ── Stage 2b: Line item extraction ───────────────────────────────────────
    logger.info("Extracting line items...", stage="2b")
    line_items = extract_line_items(sheets, scope, config, client, logger)
    logger.info(f"  {len(line_items)} line item(s) extracted", stage="2b")

    if not line_items:
        logger.error("No line items extracted — check funds flow structure", stage="2b")
        sys.exit(1)

    # ── Stage 3: Parse + extract documents ───────────────────────────────────
    logger.info("Loading support documents...", stage="3")
    cache_path = output_dir / "documents_cache.json"
    documents = load_all_documents(documents_dir, config, client, logger, cache_path)
    logger.info(f"  {len(documents)} document(s) loaded", stage="3")

    # ── Stage 4: Match ────────────────────────────────────────────────────────
    logger.info("Matching line items to documents...", stage="4")
    results = match_all(line_items, documents, config, client, logger)

    # ── Stage 5: Classify exceptions ─────────────────────────────────────────
    logger.info("Classifying exceptions...", stage="5")
    exceptions = classify_exceptions(results, documents, scope)
    logger.info(f"  {len(exceptions)} exception(s) identified", stage="5")

    # ── Stage 6c: Rename and index documents ─────────────────────────────────
    logger.info("Renaming and indexing support documents...", stage="6c")
    rename_and_index(results, documents, output_dir, logger)

    # ── Stage 6: Write outputs ────────────────────────────────────────────────
    logger.info("Writing outputs...", stage="6")
    json_path  = output_dir / "index.json"
    # Annotate the CLIENT's file in-place; save as funds_flow_indexed.xlsx
    excel_path = output_dir / "funds_flow_indexed.xlsx"

    index_data = write_index(results, exceptions, scope, config, json_path)
    summary    = index_data["summary"]

    annotate_client_file(
        source_path=funds_flow_path,
        results=results,
        exceptions=exceptions,
        scope=scope,
        all_documents=documents,
        config=config,
        summary=summary,
        output_path=excel_path,
    )

    elapsed = round(time.time() - t_start, 1)
    logger.info(f"Done in {elapsed}s — outputs at {output_dir}", stage="DONE")
    logger.close()

    # ── Stage 7: Print summary ────────────────────────────────────────────────
    _print_summary(results, exceptions, summary, elapsed)


def _print_summary(results, exceptions, summary, elapsed):
    print()
    print("=" * 72)
    print(f"  FUNDS FLOW INDEX — RESULTS")
    print("=" * 72)
    print(f"  {'Ref / Row':<22} {'Description':<36} {'Amount':>10}  Status")
    print(f"  {'-'*22} {'-'*36} {'-'*10}  {'-'*12}")
    for r in results:
        item = r.line_item
        row_id = item.source_row.split(":")[-1] if ":" in item.source_row else item.source_row
        desc = item.description[:35]
        amt  = f"${item.total_amount:,.0f}" if item.total_amount else "—"
        status = ("matched"   if r.status == "matched"   else
                  "partial"   if r.status == "partial"   else
                  "exception" if r.status == "exception" else
                  "MISSING")
        flag = " [" + ", ".join(r.exception_flags[:2]) + "]" if r.exception_flags else ""
        print(f"  {row_id:<22} {desc:<36} {amt:>10}  {status}{flag}")

    print()
    print(f"  Total:          {summary['total_line_items']}")
    print(f"  Matched:        {summary['matched']}")
    print(f"  Partial:        {summary['partial']}")
    print(f"  Missing:        {summary['missing']}")
    print(f"  Exceptions:     {summary['exceptions']}")
    print(f"  Unsupported $:  ${summary['total_unsupported_amount']:,.0f}")
    if summary.get('match_confidence_avg'):
        print(f"  Avg confidence: {summary['match_confidence_avg']:.0%}")
    print(f"  Completed in:   {elapsed}s")
    print("=" * 72)
    if exceptions:
        print(f"\n  EXCEPTIONS ({len(exceptions)} total):")
        for exc in exceptions[:10]:
            print(f"   [{exc.severity:6s}] {exc.exception_type}: {exc.description[:60]}")
        if len(exceptions) > 10:
            print(f"   ... and {len(exceptions) - 10} more — see Excel Exceptions sheet")
    print()


def _parse_args():
    p = argparse.ArgumentParser(description="Funds Flow Indexer — Production Agent")
    p.add_argument("--deal",         required=True,  help="Deal name")
    p.add_argument("--closing-date", required=True,  help="YYYY-MM-DD")
    p.add_argument("--client-role",  required=True,  choices=["buyer", "seller", "both"])
    p.add_argument("--funds-flow",   required=True,  help="Path to Excel funds flow")
    p.add_argument("--documents",    required=True,  help="Path to documents folder")
    p.add_argument("--output",       required=True,  help="Output directory")
    p.add_argument("--fund",         action="append", default=[],
                   help="Fund allocation: 'Fund I=0.60' (repeat for each fund)")
    p.add_argument("--model",        default="claude-opus-4-6")
    p.add_argument("--threshold",    type=float, default=0.75,
                   help="Minimum confidence to auto-confirm a match (0–1)")
    return p.parse_args()


def _parse_allocations(fund_args: list[str]) -> dict:
    allocs = {}
    for arg in fund_args:
        if "=" in arg:
            name, pct = arg.split("=", 1)
            try:
                allocs[name.strip()] = float(pct.strip())
            except ValueError:
                pass
    return allocs


if __name__ == "__main__":
    main()
