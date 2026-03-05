# Funds Flow Audit Tool

See `~/.claude/commands/index-funds-flow.md` for the full skill definition.

## Quick reference

1. Drop Excel + PDFs into `input/`
2. Type `/index-funds-flow`

Claude Code will:
- Stage files to `deals/<deal-name>/`
- Extract line items from the Excel
- Extract text from all PDFs
- Match documents to line items (no API key needed)
- Write `index.json`, annotated Excel, and FF-numbered document copies

Outputs land in `deals/<deal-name>/run_output/`.

## Test with Project Meridian

```bash
cp "Sample data/"* input/
```

Then `/index-funds-flow`.

## Rebuild demo data

```bash
.venv/bin/python3 archive/generate_v2_deal.py      # Project Titan PDFs
.venv/bin/python3 build_funds_flow.py               # Project Titan Excel
.venv/bin/python3 archive/generate_project_meridian.py  # Project Meridian sample data
```
