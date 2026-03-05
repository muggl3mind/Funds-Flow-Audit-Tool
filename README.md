# Funds Flow Audit Tool

A deal audit tool for private equity transactions. Given a client's funds flow Excel and a folder of supporting PDFs, it automatically matches every cost line item to its source document, annotates the workpaper, renames documents with FF-numbered references, and produces clean deliverables.

Built to run inside [Claude Code](https://docs.anthropic.com/en/docs/claude-code) using the `/index-funds-flow` command.

## What it does

```
 input/                              deals/<deal-name>/run_output/
 ├── funds_flow.xlsx        ──►     ├── funds_flow_indexed.xlsx   (annotated workpaper)
 ├── invoice_001.pdf                ├── index.json                (machine-readable index)
 ├── invoice_002.pdf                ├── documents_indexed/
 └── ...                            │   ├── FF01 - Vendor - INV-001.pdf
                                    │   ├── FF02 - Vendor - INV-002.pdf
                                    │   └── UNMATCHED/
                                    │       └── orphan_doc.pdf
                                    └── (JE tab included in workpaper)
```

For each line item in the funds flow, the indexer:

1. **Parses** the Excel workbook and detects which tabs are in scope (skips seller/wire/summary)
2. **Extracts** line items from each in-scope tab (vendor, amount, fund allocation) via LLM
3. **Parses** every PDF in the documents folder and extracts billing details (vendor, invoice number, date, amounts) via LLM with caching
4. **Matches** documents to line items using pre-filtering (amount + vendor overlap) then LLM reasoning, and assigns GL accounts from `chart_of_accounts.json`
5. **Classifies** exceptions — missing docs, partial support, amount mismatches, orphan docs
6. **Writes** an annotated Excel workpaper with audit columns, an Audit Summary sheet, and PDF snapshots
7. **Renames** matched documents with FF-numbered prefixes for clean filing

### Match statuses

| Status | Meaning |
|--------|---------|
| MATCHED | Document found, amount agrees exactly |
| CUMULATIVE | Multiple invoices from the same vendor cover the line item (e.g. interim + final billing) |
| PARTIAL | Document found but amount is less than the funds flow amount |
| MISSING | No supporting document found |

Unmatched documents (orphans that don't tie to any line item) are moved to `UNMATCHED/`.

## Setup

**Requirements:** macOS or Linux, Python 3.10+

### 1. Install Claude Code

Claude Code is a command-line tool from Anthropic. Install it once:

```bash
npm install -g @anthropic-ai/claude-code
```

> If you don't have `npm`, install [Node.js](https://nodejs.org/) first (the LTS version).
> You'll need an Anthropic API key — Claude Code will prompt you to log in on first launch.

### 2. Set up the project

```bash
# Clone the repo
git clone https://github.com/muggl3mind/Funds-Flow-Audit-Tool.git
cd Funds-Flow-Audit-Tool

# Create virtual environment and install dependencies
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

For PDF snapshot tabs in the annotated Excel (optional — `output/snapshot_tabs.py` skips gracefully if missing):
```bash
brew install poppler    # macOS — required by pdf2image
```

## Usage

### Running the indexer

1. **Drop your files** into the `input/` folder — the funds flow Excel and all supporting PDFs
2. **Launch Claude Code** by opening a terminal in the project folder and typing:
   ```bash
   claude
   ```
3. **Run the indexer** by typing this command inside Claude Code:
   ```
   /index-funds-flow
   ```

Claude Code will stage the files, match every line item to its supporting document, and print a summary table. Outputs land in `deals/<deal-name>/run_output/`.

### Standalone scripts

```bash
# Stage files from input/ to deals/<deal-name>/
.venv/bin/python3 run.py

# Or scaffold an empty deal folder
python new_deal.py --deal "Acme Acquisition" --closing-date 2026-07-15 --client-role buyer --template
```

## Project structure

```
Funds-Flow-Audit-Tool/
├── run.py                          # Staging: input/ → deals/<slug>/
├── new_deal.py                     # Scaffold empty deal folders
├── chart_of_accounts.json          # GL account reference for classification
├── requirements.txt
├── process_diagram.html            # Visual pipeline diagram (open in browser)
│
├── agent/                          # Core pipeline
│   ├── main.py                     # Full agent pipeline (Stages 1-7)
│   ├── config.py                   # DealConfig dataclass
│   ├── write_outputs.py            # Standalone orchestrator (index.json → outputs)
│   │
│   ├── parsers/                    # Stage 1 + 3: Raw text extraction
│   │   ├── excel_parser.py         # Parse Excel sheets into raw cell data
│   │   ├── pdf_parser.py           # Extract text from PDFs
│   │   └── email_parser.py         # Parse email-style PDF documents
│   │
│   ├── normalizers/                # Stages 2-3: LLM-driven normalization
│   │   ├── funds_flow_normalizer.py# Tab scope detection + line item extraction
│   │   └── document_normalizer.py  # Document billing extraction (parallel, cached)
│   │
│   ├── matcher/                    # Stage 4: Matching
│   │   ├── llm_matcher.py          # LLM-driven matching + GL classification
│   │   └── scoring.py              # Deterministic confidence scoring
│   │
│   ├── exceptions/                 # Stage 5: Post-match analysis
│   │   └── exception_classifier.py # Exception classification
│   │
│   ├── output/                     # Stage 6: Deliverables
│   │   ├── excel_writer.py         # Annotate client Excel with audit columns
│   │   ├── audit_summary.py        # Audit Summary sheet (scoreboard + exceptions)
│   │   ├── workpaper_annotator.py  # Workpaper annotation (write_outputs.py path)
│   │   ├── journal_entry_tab.py    # Journal Entry tab builder
│   │   ├── snapshot_tabs.py        # PDF snapshot tabs (optional, needs poppler)
│   │   ├── document_renamer.py     # FF-numbered document copies
│   │   ├── json_writer.py          # index.json output
│   │   └── styles.py               # Shared Excel styles
│   │
│   └── utils/                      # Shared utilities
│       ├── claude_client.py        # Anthropic API wrapper (retry, JSON extraction)
│       ├── logging_utils.py        # Structured JSON run logger
│       └── amount_utils.py         # Amount parsing/formatting
│
├── .claude/
│   └── commands/
│       └── index-funds-flow.md     # Claude Code skill definition
│
├── input/                          # Drop files here to index
├── deals/                          # Indexed deals (one folder per deal)
└── Sample data/                    # Pre-built demo files
```

## Try it out

Sample data is included so you can test immediately:

```bash
cp "Sample data/"* input/
claude
> /index-funds-flow
```

This runs **Project Meridian** — a fictional PE acquisition with 12 line items, cumulative billing, partial invoices, missing documents, and orphan docs.

## Customization

### GL accounts (chart of accounts)

The file `chart_of_accounts.json` controls which GL accounts the indexer uses to classify line items. To update it, just ask Claude in plain English:

```
> Add a new account for Insurance costs with code 7100
> Rename "Banking & Advisory Fees" to "Investment Banking Fees"
> Remove the Commercial Due Diligence account
> Replace the whole chart of accounts with our firm's GL codes
```

Claude will update the file for you. No need to edit JSON by hand.

The indexer uses judgment-based classification (not keyword rules), so account names should be descriptive of the expense type.

### Fund allocations

Fund allocation percentages (e.g. Fund I 55% / Fund II 45%) are read from the `Sources & Uses` tab of the client's funds flow Excel. No configuration needed.

### New deal scaffold

```bash
python new_deal.py --deal "Deal Name" --closing-date 2026-12-31 --client-role buyer --template
```

Creates `deals/<slug>/` with the folder structure and an optional blank funds flow template.

## How matching works

The indexer uses a two-phase matching approach:

**Phase 1 — Pre-filter** (`matcher/scoring.py`): Narrows candidates using deterministic scoring:
1. **Amount proximity** — document amount within 20% of the line item total
2. **Vendor overlap** — Jaccard token similarity on vendor names (min 15% threshold)
3. Falls back to top 10 candidates when pre-filter finds nothing

**Phase 2 — LLM reasoning** (`matcher/llm_matcher.py`): Claude evaluates the shortlisted candidates using:
1. **Reference match** — document text contains a reference number from the line item's notes
2. **Vendor match** — document text contains the vendor name from the line item description
3. **Amount match** — document total agrees with the line item amount

Edge cases handled automatically:
- **Cumulative billing** — when a later invoice supersedes an earlier one (e.g. $280K interim → $700K cumulative), both FF lines are linked
- **Partial support** — when a document covers only part of the amount (e.g. HSR filing fee without the separate SEC notification)
- **Email invoices** — PDFs containing email correspondence are tagged as `document_type: "email"`
- **Orphan documents** — cancelled engagements, subscriptions, or duplicates that don't match any line item

## License

MIT
