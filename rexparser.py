import argparse
import re
from pathlib import Path

import pdfplumber
from openpyxl import Workbook


COLUMNS = [
    "Invoice No.",
    "Date",
    "Booking ID",
    "Time",
    "Route",
    "Passenger Name",
    "Customer Ref",
    "Status / Notes",
    "Gross Price (AUD)",
    "Nett Price (AUD)",
]

DEFAULT_SOURCE_DIR = Path(
    r"C:\Users\Sonam.Tenzin\OneDrive - Programmed\PFM Analytics\Ferry Invoice Recon"
    r"\The reconning\REx\Source"
)
DEFAULT_OUTPUT_XLSX = DEFAULT_SOURCE_DIR.parent / "rex_extract.xlsx"

DATE_ROW_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{4})\s+"
    r"(?P<booking_id>\d+)\s+"
    r"(?P<time>\d{2}:\d{2})-(?P<route>[A-Z0-9-]+)\s+"
    r"(?P<middle>.*?)\s+"
    r"(?P<gross>-?\d+(?:\.\d{2})?)\s+"
    r"(?P<nett>-?\d+(?:\.\d{2})?)$"
)
INVOICE_RE = re.compile(r"Invoice no\.\s+([A-Z0-9]+)", re.IGNORECASE)
CUSTOMER_REF_RE = re.compile(r"\b\d{7}\b")


def is_noise_line(line: str) -> bool:
    if not line:
        return True

    prefixes = (
        "Page ",
        "Invoice queries",
        "T ",
        "F ",
        "E ",
        "Wadjemup Trading",
        "ABN ",
        "as trustee",
        "Trading as",
        "PO Box",
        "FREMANTLE",
        "Tax Invoice",
        "Invoice To",
        "Programmed Facility Management",
        "10 Workshop Road",
        "Rottnest Island",
        "Invoice Details",
        "Invoice no.",
        "Invoice Issue Date",
        "Payment terms",
        "Details Gross Price",
        "GST GST GST",
        "Total ",
        "Direct Deposit Payment Details",
        "Payment Reference Bank:",
        "Bank:",
        "Please quote your customer number",
        "as the payment reference",
        "BSB:",
        "email your remittance",
        "Account No:",
        "Total invoice amount",
    )
    return line.startswith(prefixes)


def split_middle(middle: str) -> tuple[str, str, str]:
    refs = list(CUSTOMER_REF_RE.finditer(middle))
    if refs:
        ref = refs[-1]
        passenger = middle[: ref.start()].strip()
        customer_ref = ref.group(0)
        status = middle[ref.end() :].strip()
        return passenger, customer_ref, status

    return middle.strip(), "", ""


def append_note(row: dict[str, str], note: str) -> None:
    note = note.strip()
    if not note:
        return

    if not row["Customer Ref"]:
        ref_match = CUSTOMER_REF_RE.match(note)
        if ref_match:
            row["Customer Ref"] = ref_match.group(0)
            note = note[ref_match.end() :].strip()
            if not note:
                return

    if row["Status / Notes"]:
        row["Status / Notes"] = f"{row['Status / Notes']} {note}"
    else:
        row["Status / Notes"] = note


def extract_line_items(pdf_path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    invoice_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""

            if not invoice_no:
                invoice_match = INVOICE_RE.search(text)
                if invoice_match:
                    invoice_no = invoice_match.group(1)

            for raw_line in text.splitlines():
                line = " ".join(raw_line.split())
                match = DATE_ROW_RE.match(line)

                if match:
                    passenger, customer_ref, status = split_middle(match.group("middle"))
                    rows.append(
                        {
                            "Invoice No.": invoice_no,
                            "Date": match.group("date"),
                            "Booking ID": match.group("booking_id"),
                            "Time": match.group("time"),
                            "Route": match.group("route"),
                            "Passenger Name": passenger,
                            "Customer Ref": customer_ref,
                            "Status / Notes": status,
                            "Gross Price (AUD)": match.group("gross"),
                            "Nett Price (AUD)": match.group("nett"),
                        }
                    )
                    continue

                if rows and not is_noise_line(line):
                    append_note(rows[-1], line)

    return rows


def write_xlsx(rows: list[dict[str, str]], xlsx_path: Path) -> None:
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Line Items"
    worksheet.append(COLUMNS)

    for row in rows:
        worksheet.append([row[column] for column in COLUMNS])

    workbook.save(xlsx_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract Rottnest Express sales invoice line items to Excel."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=str(DEFAULT_SOURCE_DIR),
        help="Path to a PDF file or a folder containing PDF files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=str(DEFAULT_OUTPUT_XLSX),
        help="Output XLSX path.",
    )
    args = parser.parse_args()

    input_path = Path(args.input_path)
    output_path = Path(args.output)

    if input_path.is_dir():
        pdf_paths = sorted(
            [path for path in input_path.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"]
        )
        if not pdf_paths:
            raise FileNotFoundError(f"No PDF files found in {input_path}")
    else:
        pdf_paths = [input_path]

    rows: list[dict[str, str]] = []
    for pdf_path in pdf_paths:
        rows.extend(extract_line_items(pdf_path))

    write_xlsx(rows, output_path)
    print(f"Extracted {len(rows)} line items to {output_path}")


if __name__ == "__main__":
    main()







