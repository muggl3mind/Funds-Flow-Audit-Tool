# Funds Flow Audit Tool

PE deal audit tool: given a client funds flow Excel + folder of support PDFs, automatically match every cost line item to its support doc, annotate the workpaper in place (auditor-style), rename docs with FF-numbered references, and produce clean deliverables.

## Key Files

| File | Role |
|------|------|
| `run.py` | Staging script — moves files from `input/` to `deals/<slug>/` |
| `new_deal.py` | Scaffold a new deal folder: `python new_deal.py --deal "Name" --closing-date YYYY-MM-DD --client-role buyer [--template]` |
| `agent/extract_funds_flow.py` | Parse Excel — extract line items from in-scope tabs |
| `agent/extract_documents.py` | Parse PDFs — extract text from all support documents |
| `agent/write_outputs.py` | Output orchestrator — reads `index.json`, delegates to sub-modules |
| `agent/output/workpaper_annotator.py` | Annotate client Excel with audit columns |
| `agent/output/journal_entry_tab.py` | Journal Entry tab builder |
| `agent/output/snapshot_tabs.py` | PDF snapshot tabs (optional, requires poppler) |
| `agent/output/styles.py` | Shared Excel styles for all output modules |
| `chart_of_accounts.json` | Reference GL accounts for line item classification |
| `.claude/commands/index-funds-flow.md` | Claude Code skill for `/index-funds-flow` |

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

## Skill Pipeline (`/index-funds-flow`)

The entire pipeline runs inside Claude Code via the `/index-funds-flow` skill:

1. **Stage** — `run.py` moves files from `input/` to `deals/<slug>/`
2. **Parse Excel** — `extract_funds_flow.py` extracts line items from in-scope tabs
3. **Parse Documents** — `extract_documents.py` extracts text from all PDFs
4. **Match & Classify** — Claude Code matches line items to documents and assigns GL accounts (in-context reasoning, no external API calls)
5. **Write `index.json`** — Claude Code writes the structured index
6. **Write Outputs** — `write_outputs.py` generates the annotated workpaper, JE tab, snapshots, and FF-numbered document copies
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
