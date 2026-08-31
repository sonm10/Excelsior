from pathlib import Path
import re
import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

# =========================================================
# FOLDER
# =========================================================

pdf_folder = Path(
    r"C:\Users\Sonam.Tenzin\OneDrive - Programmed\PFM Analytics\Ferry Invoice Recon\The reconning\FF\Sources"
)

output_excel = pdf_folder / "Combined_Ferry_InvoicesFF.xlsx"

# =========================================================
# REGEX
# =========================================================

amount_pattern = re.compile(
    r"(-?\$?\d+\.\d{2})\s+(-?\$?\d+\.\d{2})\s+(-?\$?\d+\.\d{2})"
)

header_pattern = re.compile(
    r"(\d{2}/\d{2}/\d{4})\s+Your Ref\s+([A-Z0-9]+)\s+(.*)"
)

# =========================================================
# STORAGE
# =========================================================

all_records = []

# =========================================================
# PROCESS FILES
# =========================================================

for pdf_path in pdf_folder.glob("*.pdf"):

    print(f"\nProcessing {pdf_path.name}")

    full_text = ""

    try:

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text += "\n" + text

        # -------------------------------------------------
        # CLEANING
        # -------------------------------------------------

        text = full_text

        # remove weird quotes
        text = text.replace("'", " ")

        # remove excessive whitespace
        text = re.sub(r"\s+", " ", text)

        # split by each transaction
        chunks = re.split(r"(?=\d{2}/\d{2}/\d{4}\s+Your Ref)", text)

        invoice_total = 0

        for chunk in chunks:

            chunk = chunk.strip()

            if "Your Ref" not in chunk:
                continue

            # -------------------------------------------------
            # HEADER
            # -------------------------------------------------

            h = header_pattern.search(chunk)

            if not h:
                continue

            date, ref, remaining = h.groups()

            # -------------------------------------------------
            # AMOUNTS
            # -------------------------------------------------

            amounts = amount_pattern.findall(chunk)

            if not amounts:
                continue

            # use LAST amount triplet in chunk
            ex_gst, gst, inc_gst = amounts[-1]

            def clean_amount(v):
                return float(v.replace("$", ""))

            ex_gst = clean_amount(ex_gst)
            gst = clean_amount(gst)
            inc_gst = clean_amount(inc_gst)

            invoice_total += inc_gst

            # -------------------------------------------------
            # DESCRIPTION
            # -------------------------------------------------

            desc = chunk

            # remove amounts from description
            desc = re.sub(amount_pattern, "", desc)

            # remove header
            desc = re.sub(
                r"^\d{2}/\d{2}/\d{4}\s+Your Ref\s+[A-Z0-9]+\s*",
                "",
                desc
            )

            desc = desc.strip()

            # -------------------------------------------------
            # SAVE
            # -------------------------------------------------

            all_records.append({
                "Invoice File": pdf_path.name,
                "Date": date,
                "Reference": ref,
                "Description": desc,
                "Amount Ex GST": ex_gst,
                "GST": gst,
                "Amount Inc GST": inc_gst
            })

        print(f"Extracted Invoice Total: {invoice_total:,.2f}")

    except Exception as e:
        print(f"FAILED: {pdf_path.name}")
        print(e)

# =========================================================
# DATAFRAME
# =========================================================

df = pd.DataFrame(all_records)

# =========================================================
# EXPORT EXCEL
# =========================================================

wb = Workbook()
ws = wb.active
ws.title = "Invoice Data"

ws["A1"] = "Combined Ferry Invoice Extraction"
ws["A1"].font = Font(bold=True, size=14)

for row in dataframe_to_rows(df, index=False, header=True):
    ws.append(row)

# bold headers
for cell in ws[3]:
    cell.font = Font(bold=True)

# auto width
for col in ws.columns:

    max_len = 0

    try:
        col_letter = col[0].column_letter
    except:
        continue

    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))

    ws.column_dimensions[col_letter].width = min(max_len + 2, 80)

wb.save(output_excel)

print("\nDONE")
print(f"Rows: {len(df)}")
print(f"Saved: {output_excel}")