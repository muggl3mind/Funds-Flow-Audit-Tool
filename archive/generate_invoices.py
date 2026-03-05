"""
Generate sample invoice PDFs and email PDFs for the Funds Flow Audit Tool demo.

Documents created (7 of 10 line items — INV-004, INV-006, INV-007 intentionally omitted):
  INV-001  Legal Counsel - Harper & Whitfield LLP          $750,000
  INV-002  Accounting & Audit Fees - Ashford LLP        $250,000
  INV-003  Banking Advisory - Meridian Securities       $1,500,000
  INV-004  Due Diligence - Pinnacle Advisory Group                $300,000  ← MISSING (no file)
  INV-005  Environmental Assessment - Greenleaf Consulting           $75,000
  INV-006  Tax Advisory - Sterling & Associates LLP     $150,000  ← MISSING (no file)
  INV-007  Regulatory Filing Fees - HSR/SEC               $30,000  ← MISSING (no file)
  INV-008  Miscellaneous Closing Costs                    $15,000
  EMAIL-001 IT Due Diligence - Crossbridge Partners           $85,000
  EMAIL-002 Travel & Expenses - Various                   $45,000
"""

from fpdf import FPDF
import os

OUT = os.path.join(os.path.dirname(__file__), "sample_data", "invoices")
os.makedirs(OUT, exist_ok=True)


def base_invoice(pdf: FPDF, firm: str, address: str, invoice_no: str, date: str):
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, firm, ln=True)
    pdf.set_font("Helvetica", "", 10)
    for line in address:
        pdf.cell(0, 5, line, ln=True)
    pdf.ln(8)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "INVOICE", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(60, 6, f"Invoice No: {invoice_no}")
    pdf.cell(0, 6, f"Date: {date}", ln=True)
    pdf.cell(60, 6, "Bill To:")
    pdf.cell(0, 6, "Acme Capital Partners", ln=True)
    pdf.cell(60, 6, "")
    pdf.cell(0, 6, "Attn: Fund Finance / Accounting", ln=True)
    pdf.ln(6)


def line_items_table(pdf: FPDF, items: list[tuple], total: str):
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(220, 220, 220)
    pdf.cell(120, 7, "Description", border=1, fill=True)
    pdf.cell(60, 7, "Amount (USD)", border=1, fill=True, align="R", ln=True)
    pdf.set_font("Helvetica", "", 10)
    for desc, amt in items:
        pdf.cell(120, 7, desc, border=1)
        pdf.cell(60, 7, amt, border=1, align="R", ln=True)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(120, 7, "Total Due", border=1)
    pdf.cell(60, 7, total, border=1, align="R", ln=True)


def footer_note(pdf: FPDF, note: str):
    pdf.ln(8)
    pdf.set_font("Helvetica", "I", 9)
    pdf.multi_cell(0, 5, note)


# ── INV-001: Harper & Whitfield ─────────────────────────────────────────────────
pdf = FPDF()
base_invoice(pdf, "Harper & Whitfield LLP",
             ["742 Madison Avenue", "New York, NY 10065", "Tel: (212) 555-0147"],
             "HW-2026-03142", "March 3, 2026")
line_items_table(pdf, [
    ("Deal Counsel - Project Titan Acquisition", "$500,000.00"),
    ("Negotiation & Purchase Agreement Drafting", "$175,000.00"),
    ("Closing Deliverables & Execution", "$75,000.00"),
], "$750,000.00")
footer_note(pdf, "Re: Project Titan | Engagement Letter dated Jan 15, 2026\n"
            "Payment due within 30 days. Wire instructions on file.")
pdf.output(os.path.join(OUT, "INV-001_Harper_Whitfield.pdf"))
print("Created INV-001_Harper_Whitfield.pdf")


# ── INV-002: Ashford ─────────────────────────────────────────────────────────
pdf = FPDF()
base_invoice(pdf, "Ashford LLP",
             ["500 Fifth Avenue, 14th Floor", "New York, NY 10110", "Tel: (212) 555-0238"],
             "AF-2026-00891", "February 28, 2026")
line_items_table(pdf, [
    ("Quality of Earnings Analysis - Project Titan", "$175,000.00"),
    ("Audit Support & Financial Due Diligence", "$55,000.00"),
    ("Working Capital Peg Analysis", "$20,000.00"),
], "$250,000.00")
footer_note(pdf, "Re: Project Titan | SOW dated Dec 10, 2025\n"
            "Expenses are billed at cost. Payment terms: net 30.")
pdf.output(os.path.join(OUT, "INV-002_Ashford.pdf"))
print("Created INV-002_Ashford.pdf")


# ── INV-003: Meridian Securities ────────────────────────────────────────────────────
pdf = FPDF()
base_invoice(pdf, "Meridian Securities LLC",
             ["350 Park Avenue, 22nd Floor", "New York, NY 10022", "Tel: (212) 555-0391"],
             "MS-2026-TX-0047", "March 3, 2026")
line_items_table(pdf, [
    ("M&A Advisory Fee - Project Titan Acquisition", "$1,200,000.00"),
    ("Fairness Opinion", "$200,000.00"),
    ("Financing Advisory & Capital Markets Support", "$100,000.00"),
], "$1,500,000.00")
footer_note(pdf, "Re: Project Titan | Engagement Letter dated Nov 1, 2025\n"
            "Fee contingent on closing. Closing occurred March 3, 2026. Payment due immediately.")
pdf.output(os.path.join(OUT, "INV-003_Meridian_Securities.pdf"))
print("Created INV-003_Meridian_Securities.pdf")


# ── INV-004: Pinnacle Advisory Group ── INTENTIONALLY OMITTED ─────────────────────────
# Due Diligence - $300,000 — no file created
print("SKIPPED INV-004_Pinnacle_Advisory.pdf (intentional gap)")


# ── INV-005: Greenleaf Consulting ───────────────────────────────────────────────────────
pdf = FPDF()
base_invoice(pdf, "Greenleaf Consulting",
             ["275 Seventh Avenue, Suite 800", "New York, NY 10001", "Tel: (212) 555-0472"],
             "GL-2026-1847", "February 20, 2026")
line_items_table(pdf, [
    ("Phase I Environmental Site Assessment", "$40,000.00"),
    ("Phase II ESG Risk Assessment", "$25,000.00"),
    ("Report Preparation & Management Summary", "$10,000.00"),
], "$75,000.00")
footer_note(pdf, "Re: Project Titan | Scope of Work dated Jan 5, 2026\n"
            "Final report delivered Feb 18, 2026. Payment net 30.")
pdf.output(os.path.join(OUT, "INV-005_Greenleaf_Consulting.pdf"))
print("Created INV-005_Greenleaf_Consulting.pdf")


# ── INV-006: PwC ── INTENTIONALLY OMITTED ────────────────────────────────────
# Tax Advisory - $150,000 — no file created
print("SKIPPED INV-006_Sterling.pdf (intentional gap)")


# ── INV-007: HSR Filing Fees ── INTENTIONALLY OMITTED ────────────────────────
# Regulatory Filing Fees - $30,000 — no file created
print("SKIPPED INV-007_HSR_Filing.pdf (intentional gap)")


# ── INV-008: Miscellaneous ────────────────────────────────────────────────────
pdf = FPDF()
base_invoice(pdf, "Acme Capital Partners - Internal Accounting",
             ["888 Park Avenue, Suite 3200", "New York, NY 10021"],
             "ACP-MISC-2026-031", "March 3, 2026")
line_items_table(pdf, [
    ("Notary & Document Authentication Fees", "$4,500.00"),
    ("Overnight Courier & Delivery (FedEx/UPS)", "$2,200.00"),
    ("Virtual Data Room - Intralinks (3 months)", "$6,300.00"),
    ("Administrative & Printing Costs", "$2,000.00"),
], "$15,000.00")
footer_note(pdf, "Re: Project Titan closing costs | Compiled from receipts on file.\n"
            "Allocated 60/40 between Fund I and Fund II per capital commitments.")
pdf.output(os.path.join(OUT, "INV-008_Misc_Closing.pdf"))
print("Created INV-008_Misc_Closing.pdf")


# ── EMAIL-001: Crossbridge Partners ───────────────────────────────────────────────
pdf = FPDF()
pdf.add_page()
pdf.set_font("Helvetica", "B", 13)
pdf.cell(0, 10, "EMAIL CORRESPONDENCE", ln=True)
pdf.set_font("Helvetica", "", 10)
pdf.set_fill_color(240, 240, 240)
pdf.cell(0, 7, "From:    d.ramirez@crossbridgepartners.com", ln=True, fill=True)
pdf.cell(0, 7, "To:      j.kim@acmecapital.com", ln=True)
pdf.cell(0, 7, "Date:    February 25, 2026  9:14 AM EST", ln=True, fill=True)
pdf.cell(0, 7, "Subject: Project Titan - IT Due Diligence Final Invoice", ln=True)
pdf.ln(6)
pdf.set_font("Helvetica", "", 10)
pdf.multi_cell(0, 6,
    "Jane,\n\n"
    "Please find below our final invoice summary for IT due diligence services "
    "rendered in connection with Project Titan.\n\n"
    "  Services Rendered:  IT Systems Architecture Review & Cybersecurity Assessment\n"
    "  Period:             December 2025 - February 2026\n"
    "  Invoice Reference:  CB-IT-2026-0312\n"
    "  Total Amount Due:   $85,000.00\n\n"
    "Breakdown:\n"
    "  - IT Infrastructure & Systems Assessment      $45,000.00\n"
    "  - Cybersecurity Vulnerability Review          $25,000.00\n"
    "  - ERP & Data Management Review               $15,000.00\n"
    "  -----------------------------------------------\n"
    "  Total                                         $85,000.00\n\n"
    "Wire instructions were provided under separate cover. Please confirm receipt "
    "and expected payment date at your earliest convenience.\n\n"
    "Best regards,\n"
    "Daniel Ramirez\n"
    "Managing Director, Technology Risk & Diligence\n"
    "Crossbridge Partners | New York"
)
pdf.output(os.path.join(OUT, "EMAIL-001_Crossbridge_Partners.pdf"))
print("Created EMAIL-001_Crossbridge_Partners.pdf")


# ── EMAIL-002: Travel & Expenses ──────────────────────────────────────────────
pdf = FPDF()
pdf.add_page()
pdf.set_font("Helvetica", "B", 13)
pdf.cell(0, 10, "EMAIL CORRESPONDENCE", ln=True)
pdf.set_font("Helvetica", "", 10)
pdf.set_fill_color(240, 240, 240)
pdf.cell(0, 7, "From:    s.patel@acmecapital.com", ln=True, fill=True)
pdf.cell(0, 7, "To:      finance@acmecapital.com", ln=True)
pdf.cell(0, 7, "Date:    March 2, 2026  4:52 PM EST", ln=True, fill=True)
pdf.cell(0, 7, "Subject: Project Titan - Deal Team T&E Summary for Closing", ln=True)
pdf.ln(6)
pdf.set_font("Helvetica", "", 10)
pdf.multi_cell(0, 6,
    "Finance Team,\n\n"
    "Attached is the consolidated T&E summary for the deal team in connection "
    "with Project Titan diligence and closing. All receipts are in Concur.\n\n"
    "  Reference:          ACP-TE-2026-TITAN\n"
    "  Total Amount:       $45,000.00\n\n"
    "Summary by category:\n"
    "  - Airfare (New York / Chicago / Dallas)       $18,500.00\n"
    "  - Hotel & Lodging (12 nights, 4 team members) $14,200.00\n"
    "  - Ground Transportation & Car Service          $6,800.00\n"
    "  - Meals & Entertainment (client-facing)        $5,500.00\n"
    "  -----------------------------------------------\n"
    "  Total                                         $45,000.00\n\n"
    "Please book the full $45,000 to Transaction Costs - T&E per the funds flow. "
    "Allocation is 60/40 Fund I / Fund II.\n\n"
    "Thanks,\n"
    "Samira Patel\n"
    "Vice President, Finance & Accounting\n"
    "Acme Capital Partners"
)
pdf.output(os.path.join(OUT, "EMAIL-002_Travel_Expenses.pdf"))
print("Created EMAIL-002_Travel_Expenses.pdf")

print("\nDone. Files in:", OUT)
print("Intentional gaps: INV-004 (Pinnacle), INV-006 (Sterling), INV-007 (HSR Filing)")
