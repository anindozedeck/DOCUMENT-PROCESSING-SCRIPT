import argparse
import csv
import os
from datetime import date

import anthropic
from docx import Document
from docx.shared import Pt, RGBColor
from dotenv import load_dotenv
from pypdf import PdfReader


def summarize_contract():
    load_dotenv()
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    reader = PdfReader("supplier_contract.pdf")
    print(f"Pages: {len(reader.pages)}")

    doc_text = "\n".join(page.extract_text() or "" for page in reader.pages)
    if not doc_text.strip():
        raise ValueError("No extractable text found; scanned PDFs require OCR.")

    response = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=512,
        messages=[
            {
                "role": "user",
                "content": (
                    "From this contract, extract the party names, payment terms, "
                    "delivery date, and any penalties. If a detail is not stated, "
                    "say 'Not specified'.\n\n"
                    f"Document:\n{doc_text[:3000]}"
                ),
            }
        ],
    )
    print(response.content[0].text)


def create_sales_report():
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)

    title = doc.add_heading("Monthly Sales Report", 0)
    title.runs[0].font.color.rgb = RGBColor(31, 78, 121)

    metadata = doc.add_paragraph()
    metadata.add_run("Prepared by: Kamau Stores | Date: ").bold = True
    metadata.add_run(date.today().strftime("%B %d, %Y"))

    doc.add_heading("Summary", level=1)
    doc.add_paragraph(
        "Total sales for November 2024: KES 847,000. Up 12% from October."
    )

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    headers = table.rows[0].cells
    headers[0].text = "Product"
    headers[1].text = "Units"
    headers[2].text = "Revenue (KES)"
    for cell in headers:
        for run in cell.paragraphs[0].runs:
            run.bold = True

    data = [
        ("Maize flour", 420, 336000),
        ("Rice", 180, 324000),
        ("Cooking oil", 210, 187000),
    ]
    for product, units, revenue in data:
        cells = table.add_row().cells
        cells[0].text = product
        cells[1].text = f"{units:,}"
        cells[2].text = f"{revenue:,}"

    doc.save("november_report.docx")
    print("Report saved: november_report.docx")


def index_pdf_folder(pdf_folder="invoices", output_csv="invoice_index.csv"):
    if not os.path.isdir(pdf_folder):
        raise FileNotFoundError(f"PDF folder not found: {pdf_folder}")

    results = []
    for fname in sorted(os.listdir(pdf_folder)):
        if not fname.lower().endswith(".pdf"):
            continue

        path = os.path.join(pdf_folder, fname)
        reader = PdfReader(path)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        results.append(
            {
                "file": fname,
                "pages": len(reader.pages),
                "chars": len(text),
                "preview": text[:100].replace("\n", " ").replace("\r", " "),
            }
        )

    fieldnames = ["file", "pages", "chars", "preview"]
    with open(output_csv, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)
    print(f"Indexed {len(results)} invoices to {output_csv}")


def main():
    parser = argparse.ArgumentParser(description="Process PDF and Word documents.")
    parser.add_argument(
        "workflow",
        nargs="?",
        choices=("documents", "batch"),
        default="documents",
        help="'documents' runs the existing contract/report flow; 'batch' indexes PDFs.",
    )
    parser.add_argument("pdf_folder", nargs="?", default="invoices")
    parser.add_argument("--output", default="invoice_index.csv")
    args = parser.parse_args()

    if args.workflow == "batch":
        index_pdf_folder(args.pdf_folder, args.output)
    else:
        summarize_contract()
        create_sales_report()


if __name__ == "__main__":
    main()