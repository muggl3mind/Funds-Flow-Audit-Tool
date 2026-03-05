"""
Parse monetary amounts from every format seen in real PE funds flows and documents.
"""
from __future__ import annotations

import re
from typing import Optional


_STRIP = re.compile(r"[\$£€¥,\s]")
_PARENS = re.compile(r"^\((.+)\)$")
_MILLIONS = re.compile(r"^([\d.]+)\s*[Mm]$")
_THOUSANDS = re.compile(r"^([\d.]+)\s*[Kk]$")
_BILLIONS = re.compile(r"^([\d.]+)\s*[Bb]$")


def parse_amount(raw_value, unit_scale: float = 1.0) -> Optional[float]:
    """
    Return a float dollar amount or None if unparseable.

    Handles:
      int / float from openpyxl (most common — just multiply by unit_scale)
      "$750,000"  "$750,000.00"  "750,000"
      "(750,000)"  "($750,000)"  negative accounting format
      "-$750,000"
      "$1.2M"  "1.2M"  "$750K"  "2.5B"
      "" / None / "—" / "N/A"

    unit_scale: multiply result by this factor.
      Use 1000.0 when openpyxl number_format contains "#,##0," (thousands scaling).
    """
    if raw_value is None:
        return None

    # Native numeric types from openpyxl
    if isinstance(raw_value, (int, float)):
        return float(raw_value) * unit_scale

    text = str(raw_value).strip()

    # Blank / placeholder
    if not text or text in ("-", "—", "N/A", "n/a", "nil", "Nil", "TBD"):
        return None

    negative = False

    # Accounting negatives: (750,000) or ($750,000)
    m = _PARENS.match(text)
    if m:
        text = m.group(1)
        negative = True

    # Dash-negative
    if text.startswith("-"):
        text = text[1:]
        negative = True

    # Strip currency symbols, commas, spaces
    text = _STRIP.sub("", text)

    # Shorthand suffixes
    m = _BILLIONS.match(text)
    if m:
        value = float(m.group(1)) * 1_000_000_000
        return (-value if negative else value) * unit_scale

    m = _MILLIONS.match(text)
    if m:
        value = float(m.group(1)) * 1_000_000
        return (-value if negative else value) * unit_scale

    m = _THOUSANDS.match(text)
    if m:
        value = float(m.group(1)) * 1_000
        return (-value if negative else value) * unit_scale

    # Plain numeric
    try:
        value = float(text)
        return (-value if negative else value) * unit_scale
    except ValueError:
        return None


def infer_unit_scale(number_format: str) -> float:
    """
    Detect Excel thousands / millions scaling from the cell number format string.
    '#,##0,'  → thousands (scale 1000)
    '#,##0,,' → millions  (scale 1_000_000)
    """
    if not number_format:
        return 1.0
    trailing_commas = len(number_format.rstrip()) - len(number_format.rstrip().rstrip(","))
    if trailing_commas == 1:
        return 1_000.0
    if trailing_commas >= 2:
        return 1_000_000.0
    return 1.0


def format_usd(amount: Optional[float]) -> str:
    if amount is None:
        return "—"
    return f"${amount:,.0f}"
