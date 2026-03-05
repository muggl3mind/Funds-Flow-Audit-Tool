# Funds Flow Audit Tool

You are the Funds Flow Audit Tool agent. Follow every step below in order.
Do not skip steps. Do not call any external API. Use only your tools (Bash, Read, Write).

---

## Step 0 — Check input and stage files

Run:
```bash
.venv/bin/python3 run.py
```

If `input/` is empty, stop and tell the user to drop their funds flow Excel and PDF documents into the `input/` folder, then run `/index-funds-flow` again.

After staging succeeds, read `last_run.json` to get the deal directory and metadata.

---

## Step 1 — Extract funds flow line items

Run:
```bash
.venv/bin/python3 agent/extract_funds_flow.py "<deal_dir>/funds_flow.xlsx" "<deal_dir>/run_output/ff_extract.json"
```

(Replace `<deal_dir>` with the value from `last_run.json`.)

Then read `<deal_dir>/run_output/ff_extract.json`.

---

## Step 2 — Extract document text

Run:
```bash
.venv/bin/python3 agent/extract_documents.py "<deal_dir>/documents" "<deal_dir>/run_output/docs_extract.json"
```

Then read `<deal_dir>/run_output/docs_extract.json`.

---

## Step 3 — Match line items to documents and classify GL accounts

Read `chart_of_accounts.json` from the project root. This is the reference list of available GL accounts.

For each line item in `ff_extract.json`:

**3a. Document matching** — find the best matching document from `docs_extract.json`:
1. **Reference match** — the document text contains a reference number found in the line item's `notes` field
2. **Vendor match** — the document text contains the vendor name from the line item description
3. **Amount match** — the document text mentions an amount within ±5% of the line item total

Assign each line item one of these statuses:
- `MATCHED` — one document found; amount agrees exactly (within $1 rounding)
- `PARTIAL` — document found but amount is less than the funds flow amount
- `CUMULATIVE` — two or more documents together add up to the line item total (same vendor, multi-invoice)
- `MISSING` — no document found or no confident match

Also identify any documents that don't match any line item → mark as `UNMATCHED`.

For cumulative billing: if two line items share a vendor and one document covers both combined totals, link the document to the larger/final line item and note the relationship.

**3b. GL account classification** — for each line item, select the most appropriate account from `chart_of_accounts.json` using judgment:
- Read the description, vendor name, document type, and context
- Do not match keywords mechanically — reason about what the expense actually is
- Examples: "SPA execution" → Legal Fees; "fairness opinion" → Banking & Advisory; "QofE analysis" → Financial Due Diligence; "HSR filing fee" → Regulatory & Filing Fees
- If genuinely ambiguous, pick the closest account and note it in the `notes` field
- Every line item must have a `gl_account_code` and `gl_account_name` — including MISSING items

---

## Step 4 — Write index.json

Write the results to `<deal_dir>/run_output/index.json` in this exact structure:

```json
{
  "deal": "<deal name>",
  "closing_date": "<YYYY-MM-DD>",
  "client_role": "<buyer|seller|both>",
  "fund_allocations": { "Fund I": 0.55, "Fund II": 0.45 },
  "indexed_at": "<today's date YYYY-MM-DD>",
  "line_items": [
    {
      "ff_ref": "FF01",
      "status": "MATCHED",
      "gl_account_code": "7010",
      "gl_account_name": "Transaction Costs — Legal Fees",
      "description": "Legal Counsel — Harrington & Cross LLP",
      "tab": "Buyer Expenses",
      "funds_flow_amount": 700000,
      "fund_i_amount": 385000,
      "fund_ii_amount": 315000,
      "document_file": "HarringtonCross_Invoice002_Cumulative_Jul2026.pdf",
      "document_vendor": "Harrington & Cross LLP",
      "document_amount": 700000,
      "amount_agrees": true,
      "document_type": "invoice",
      "notes": ""
    }
  ],
  "unmatched_documents": ["LexPro_Research_Q2_2026.pdf"],
  "summary": {
    "total_line_items": 0,
    "matched": 0,
    "partial": 0,
    "cumulative": 0,
    "missing": 0,
    "unmatched_documents": 0,
    "total_funds_flow_amount": 0,
    "total_supported_amount": 0,
    "total_unsupported_amount": 0
  }
}
```

Compute all summary totals accurately.

---

## Step 5 — Write final outputs

Run:
```bash
PYTHONPATH=. .venv/bin/python3 agent/write_outputs.py "<deal_dir>"
```

This creates:
- `<deal_dir>/run_output/funds_flow_indexed.xlsx` — annotated workpaper
- `<deal_dir>/run_output/documents_indexed/` — FF-numbered matched documents
- `<deal_dir>/run_output/documents_indexed/UNMATCHED/` — orphan documents

---

## Step 6 — Report to user

Print a clean summary table:

```
=====================================================================
  FUNDS FLOW INDEX — <Deal Name>
  Closing: <date>   Role: <BUYER/SELLER>
=====================================================================
  FF#   Description                          Amount       Status
  ---   -----------                          ------       ------
  FF01  Harrington & Cross — Due Diligence   $280,000     CUMULATIVE
  FF02  Harrington & Cross — Closing         $420,000     CUMULATIVE
  ...
  —     Vertex Strategy Group                $250,000     MISSING
=====================================================================
  Matched:    X   |   Partial: X   |   Missing: X
  Supported:  $X,XXX,XXX   |   Unsupported: $XXX,XXX
  Unmatched documents: X (see UNMATCHED/)
=====================================================================
```

---

## Notes

- Assign `ff_ref` sequentially (FF01, FF02, …) to every line item in order. MISSING items get `null`. The same FF number is used in the workpaper annotation column, the snapshot tab name, and the `documents_indexed/` filename — `write_outputs.py` writes `ff_ref` back into `index.json` automatically, so it is always consistent.
- Do not fabricate document contents. If text extraction fails for a document, mark the line item as MISSING.
- For email invoices (filename contains "Email"), set `document_type` to `"email"`.
- The `Seller Expenses` and `Wire Instructions` tabs are always skipped.
- The `Sources & Uses` tab is a summary — line items come from the detail tabs only.
