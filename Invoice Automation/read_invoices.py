import pdfplumber
import os

INVOICES_DIR = r"c:\Users\lenovo\OneDrive\AWS Automation Codes\Invoice Automation\invoices"

files = sorted([f for f in os.listdir(INVOICES_DIR) if f.startswith("Invoice_2") and f.endswith(".pdf")])

for fname in files:
    print("=" * 80)
    print("FILE:", fname)
    print("=" * 80)
    with pdfplumber.open(os.path.join(INVOICES_DIR, fname)) as pdf:
        print("TOTAL PAGES:", len(pdf.pages))
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text and text.strip():
                print(f"\n--- PAGE {i+1} ---")
                print(text)
    print()
