"""Shared Excel styles for all output modules."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

STATUS_COLORS = {
    "MATCHED":    "C6EFCE",
    "PARTIAL":    "FFEB9C",
    "CUMULATIVE": "DDEBF7",
    "MISSING":    "FFC7CE",
    "UNMATCHED":  "E2EFDA",
}

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
CELL_FONT   = Font(name="Calibri", size=10)
BOLD_FONT   = Font(name="Calibri", bold=True, size=10)

THIN_SIDE   = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
MED_BOTTOM  = Border(bottom=Side(style="medium", color="1F3864"))

USD_FMT = '"$"#,##0.00'
CHK_FMT = '"$"#,##0.00_);[Red]("$"#,##0.00)'
