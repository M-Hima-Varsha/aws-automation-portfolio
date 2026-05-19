"""
Run this script once to generate dummy AWS invoice PDFs for testing.

    pip install reportlab
    python generate_dummy_invoices.py

Three PDF files will be created inside the invoices/ folder:
    Invoice_DEMO0000001.pdf  — Dev account  (AWS services + Marketplace)
    Invoice_DEMO0000002.pdf  — Prod account (AWS services only)
    Invoice_DEMO0000003.pdf  — Sandbox account (AWS services + Late Fee)
"""

import os
import sys

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
except ImportError:
    sys.exit("reportlab is not installed. Run:  pip install reportlab")

INVOICES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoices")
os.makedirs(INVOICES_DIR, exist_ok=True)


def create_invoice(filename, lines):
    path = os.path.join(INVOICES_DIR, filename)
    c = canvas.Canvas(path, pagesize=letter)
    _, height = letter
    y = height - 50
    for line in lines:
        if y < 50:
            c.showPage()
            y = height - 50
        c.setFont("Courier", 10)
        c.drawString(40, y, line)
        y -= 14
    c.save()
    print(f"  Created : {filename}")


# ── Invoice 1 — Dev account — AWS Services + Marketplace charge ───────────────
create_invoice("Invoice_DEMO0000001.pdf", [
    "Amazon Web Services, Inc.",
    "Invoice Number: DEMO0000001",
    "Account Number:",
    "111122223333",
    "Bill to Address:",
    "Demo Organization",
    "Please include this invoice number with your payment.",
    "This invoice is for the billing period April 1 - April 30, 2026",
    "",
    "Summary",
    "Dev-Account-Alpha (111122223333) USD 520.75",
    "Charges USD 520.75",
    "",
    "Detail for Linked Account: Dev-Account-Alpha (111122223333)",
    "Dev-Account-Alpha (111122223333) USD 520.75",
    "Charges USD 520.75",
    "Amazon EC2 USD 310.50",
    "Amazon S3 USD 85.25",
    "Amazon RDS USD 75.00",
    "DataDog Agent sold by Datadog USD 50.00",
])

# ── Invoice 2 — Prod account — AWS Services only ──────────────────────────────
create_invoice("Invoice_DEMO0000002.pdf", [
    "Amazon Web Services, Inc.",
    "Invoice Number: DEMO0000002",
    "Account Number:",
    "444455556666",
    "Bill to Address:",
    "Demo Organization",
    "Please include this invoice number with your payment.",
    "This invoice is for the billing period April 1 - April 30, 2026",
    "",
    "Summary",
    "Prod-Account-Beta (444455556666) USD 1245.60",
    "Charges USD 1245.60",
    "",
    "Detail for Linked Account: Prod-Account-Beta (444455556666)",
    "Prod-Account-Beta (444455556666) USD 1245.60",
    "Charges USD 1245.60",
    "Amazon EC2 USD 780.00",
    "Amazon RDS USD 310.60",
    "AWS Lambda USD 155.00",
])

# ── Invoice 3 — Sandbox account — AWS Services + Late Fee ─────────────────────
create_invoice("Invoice_DEMO0000003.pdf", [
    "Amazon Web Services, Inc.",
    "Invoice Number: DEMO0000003",
    "Account Number:",
    "777788889999",
    "Bill to Address:",
    "Demo Organization",
    "Please include this invoice number with your payment.",
    "This invoice is for the billing period April 1 - April 30, 2026",
    "",
    "Summary",
    "Sandbox-Account-Gamma (777788889999) USD 198.40",
    "Charges USD 198.40",
    "",
    "Detail for Linked Account: Sandbox-Account-Gamma (777788889999)",
    "Sandbox-Account-Gamma (777788889999) USD 198.40",
    "Charges USD 198.40",
    "Amazon EC2 USD 120.00",
    "Amazon CloudWatch USD 48.40",
    "Amazon S3 USD 30.00",
    "",
    "Detail",
    "OCBLateFee USD 15.00",
])

print("\nAll 3 dummy invoices created in the invoices/ folder.")
print("Now run:  python invoice-automation.py")
