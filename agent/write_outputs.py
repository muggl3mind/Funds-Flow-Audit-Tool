"""
agent/write_outputs.py — Orchestrate final output generation.

Reads index.json, assigns FF numbers, then delegates to focused modules:
  - workpaper_annotator  — annotate client Excel with match-status columns
  - journal_entry_tab    — build Journal Entry tab (Fund I / Fund II)
  - snapshot_tabs        — render PDF snapshots into Excel tabs
  - document copying     — FF-numbered copies + UNMATCHED folder

Usage:
  python agent/write_outputs.py <deal_dir>
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

from agent.output.workpaper_annotator import annotate
from agent.output.journal_entry_tab import build as build_je_tab
from agent.output.snapshot_tabs import add_snapshots


def _assign_ff_numbers(index: dict) -> None:
    """Assign FF refs sequentially; MISSING items get None. Mutates index in place."""
    ff_num = 1
    for item in index["line_items"]:
        if item.get("status") != "MISSING":
            item["ff_ref"] = f"FF{ff_num:02d}"
        else:
            item["ff_ref"] = None
        ff_num += 1


def _copy_numbered_docs(deal_dir: Path, index: dict) -> tuple[int, int]:
    """Copy matched docs with FF-numbered names; move unmatched to UNMATCHED/."""
    docs_dir      = deal_dir / "documents"
    idx_dir       = deal_dir / "run_output" / "documents_indexed"
    unmatched_dir = idx_dir / "UNMATCHED"

    if idx_dir.exists():
        shutil.rmtree(idx_dir)
    idx_dir.mkdir(parents=True)
    unmatched_dir.mkdir()

    matched_files: set[str] = set()
    matched_count = 0

    for item in index["line_items"]:
        doc_file = item.get("document_file")
        ff_ref   = item.get("ff_ref")
        if not doc_file or not ff_ref or item.get("status") == "MISSING":
            continue
        src = docs_dir / doc_file
        if not src.exists():
            continue

        vendor    = (item.get("document_vendor") or "").replace("/", "-")[:30].strip()
        dest_name = f"{ff_ref} - {vendor} - {doc_file}"
        shutil.copy(str(src), str(idx_dir / dest_name))
        matched_files.add(doc_file)
        matched_count += 1
        print(f"  {ff_ref}: {doc_file}")

    unmatched_count = 0
    for pdf in sorted(docs_dir.glob("*.pdf")):
        if pdf.name not in matched_files:
            shutil.copy(str(pdf), str(unmatched_dir / pdf.name))
            print(f"  UNMATCHED: {pdf.name}")
            unmatched_count += 1

    return matched_count, unmatched_count


def main():
    if len(sys.argv) < 2:
        print("Usage: python agent/write_outputs.py <deal_dir>")
        sys.exit(1)

    deal_dir   = Path(sys.argv[1])
    index_path = deal_dir / "run_output" / "index.json"

    if not index_path.exists():
        print(f"ERROR: index.json not found at {index_path}")
        sys.exit(1)

    index = json.loads(index_path.read_text())

    # Assign FF numbers and persist back
    _assign_ff_numbers(index)
    index_path.write_text(json.dumps(index, indent=2))

    # 1. Annotate workpaper
    print("Writing annotated Excel...")
    wb, out_path, je_row_info = annotate(deal_dir, index)

    # 2. Journal Entry tab
    print("  Building Journal Entry tab...")
    build_je_tab(wb, index, je_row_info)

    # 3. PDF snapshots
    add_snapshots(wb, deal_dir, index)

    wb.save(str(out_path))
    print(f"  → {out_path.name}")

    # 4. Copy numbered documents
    print("\nCopying numbered documents...")
    matched, unmatched = _copy_numbered_docs(deal_dir, index)
    print(f"  → {matched} matched, {unmatched} unmatched")

    print("\nDone.")


if __name__ == "__main__":
    main()
