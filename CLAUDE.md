# Funds Flow Indexer

PE deal audit tool: given a client funds flow Excel + folder of support PDFs, automatically match every cost line item to its support doc, annotate the workpaper in place (auditor-style), rename docs with FF-numbered references, and produce clean deliverables.

## Key Files

| File | Role |
|------|------|
| `run.py` | Staging script — moves files from `input/` to `deals/<slug>/` |
| `new_deal.py` | Scaffold a new deal folder: `python new_deal.py --deal "Name" --closing-date YYYY-MM-DD --client-role buyer [--template]` |
| `agent/main.py` | Full production agent pipeline (Stages 1-7) |
| `agent/config.py` | `DealConfig` dataclass — all deal parameters |
| `agent/parsers/excel_parser.py` | Stage 1 — parse Excel sheets into raw cell data |
| `agent/parsers/pdf_parser.py` | Raw PDF text extraction |
| `agent/parsers/email_parser.py` | Email/PDF document text parsing |
| `agent/normalizers/funds_flow_normalizer.py` | Stage 2 — tab scope detection + line item extraction |
| `agent/normalizers/document_normalizer.py` | Stage 3 — LLM-driven document normalization (parallel, cached) |
| `agent/matcher/llm_matcher.py` | Stage 4 — LLM-driven matching + GL classification |
| `agent/matcher/scoring.py` | Deterministic confidence scoring (no LLM) |
| `agent/exceptions/exception_classifier.py` | Stage 5 — post-match exception classification |
| `agent/output/excel_writer.py` | Stage 6 — annotates client Excel with audit columns |
| `agent/output/audit_summary.py` | Audit Summary sheet (scoreboard + exceptions) |
| `agent/output/json_writer.py` | `index.json` output |
| `agent/output/document_renamer.py` | Stage 6c — renames + indexes docs |
| `agent/output/styles.py` | Shared Excel styles for all output modules |
| `agent/write_outputs.py` | Standalone orchestrator — reads `index.json`, delegates to sub-modules |
| `agent/output/workpaper_annotator.py` | Workpaper annotation (used by `write_outputs.py`) |
| `agent/output/journal_entry_tab.py` | Journal Entry tab builder |
| `agent/output/snapshot_tabs.py` | PDF snapshot tabs (optional, requires poppler) |
| `agent/utils/claude_client.py` | Anthropic API wrapper (retry, JSON extraction) |
| `agent/utils/logging_utils.py` | Structured JSON run logger |
| `agent/utils/amount_utils.py` | Amount parsing/formatting utilities |
| `chart_of_accounts.json` | Reference GL accounts for line item classification |
| `.claude/commands/index-funds-flow.md` | Claude Code command for `/index-funds-flow` |

## Running

```bash
# Use the project venv for all Python commands
.venv/bin/python3 <script>
```

### Quick start
1. Drop funds flow Excel + support PDFs into `input/`
2. Run `/index-funds-flow`
3. Outputs land in `deals/<deal-name>/run_output/`

### New deal scaffold
```bash
python new_deal.py --deal "Deal Name" --closing-date YYYY-MM-DD --client-role buyer [--template]
# Creates deals/<slug>/{documents/, run_output/, README.txt, [funds_flow_template.xlsx]}
```

## Agent Pipeline (Stages)

1. **Parse Excel** — `parsers/excel_parser.py` reads all sheets + cell values
2. **Scope & Extraction** — `normalizers/funds_flow_normalizer.py`
   - 2a: Tab scope detection (rule-based + LLM fallback)
   - 2b: Line item extraction (LLM)
3. **Document Parsing** — `normalizers/document_normalizer.py` (parallel, cached via `documents_cache.json`)
   - Uses `parsers/pdf_parser.py` + `parsers/email_parser.py` for text extraction
   - LLM extracts vendor, invoice number, date, amounts per document
4. **Matching** — `matcher/llm_matcher.py` pre-filters candidates via `matcher/scoring.py`, then LLM picks best match + assigns GL account
5. **Exception Classification** — `exceptions/exception_classifier.py`
6. **Write Outputs** — `output/excel_writer.py` (audit columns) + `output/audit_summary.py` + `output/json_writer.py`
6c. **Rename & Index Docs** — `output/document_renamer.py` — runs every time automatically
7. **Print Summary** — console results table

## Matching Rules

- **MATCHED** — one document found; amount agrees exactly (within rounding)
- **PARTIAL** — document found but amount is less than the funds flow amount
- **CUMULATIVE** — two or more documents/invoices together cover the line item total (same vendor, multi-invoice)
- **MISSING** — no document found or no confident match
- **UNMATCHED** — document doesn't match any line item (orphan)
- Skip `Seller Expenses` and `Wire Instructions` tabs
- `Sources & Uses` is a summary tab — line items come from detail tabs only
- For email invoices (filename contains "Email"), set `document_type` to `"email"`

## Journal Entry Rules

- JE dates MUST use the invoice date from the supporting document, NOT the deal closing date
- This applies to all JE generation: the JE tab in the indexed workpaper and `/finance:journal-entry-prep`

## Sample Data

All sample data uses fictional company names. Test with:
```bash
cp "Sample data/"* input/
# Then /index-funds-flow
```
