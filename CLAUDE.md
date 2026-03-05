# Funds Flow Indexer

PE deal audit tool: given a client funds flow Excel + folder of support PDFs, automatically match every cost line item to its support doc, annotate the workpaper in place (auditor-style), rename docs with FF-numbered references, and produce clean deliverables.

## Key Files

| File | Role |
|------|------|
| `run.py` | Staging script — moves files from `input/` to `deals/<slug>/` |
| `new_deal.py` | Scaffold a new deal folder: `python new_deal.py --deal "Name" --closing-date YYYY-MM-DD --client-role buyer [--template]` |
| `agent/extract_funds_flow.py` | Extract line items from funds flow Excel |
| `agent/extract_documents.py` | Extract text from PDF documents |
| `agent/write_outputs.py` | Write annotated Excel, FF-numbered documents, and JE tab |
| `agent/main.py` | Full production agent pipeline (Stages 1-6c) |
| `agent/output/document_renamer.py` | Stage 6c — renames + indexes docs |
| `agent/output/excel_writer.py` | Stage 6 — annotates client Excel (no snapshots) |
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

1. Parse Excel (`excel_parser.py`)
2. Tab scope detection + line item extraction
3. Load + parse documents (`document_normalizer.py`, parallel, cached)
4. Match line items to documents
5. Exception classification
6. Write outputs (`excel_writer.py` + `json_writer.py`)
6c. Rename + index docs (`document_renamer.py`) — runs every time automatically

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
