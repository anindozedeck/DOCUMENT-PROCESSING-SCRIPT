# Document Processing Script

A small Python demo for extracting information from PDFs, generating a Word
sales report, and indexing batches of PDF invoices into a CSV file.

## Features

- **Contract summary:** extracts text from `supplier_contract.pdf` and asks
  Anthropic to identify the parties, payment terms, delivery date, and
  penalties.
- **Sales report:** creates `november_report.docx` with sample sales data.
- **PDF batch index:** scans a folder for PDF files and writes each file's page
  count, extracted-text character count, and a short preview to a CSV.

## Requirements

- Python with `pip`
- An Anthropic API key for the default contract-summary workflow

Install the required packages:

```powershell
python -m pip install anthropic python-dotenv pypdf python-docx
```

For the contract-summary workflow, create a `.env` file in the project folder:

```text
ANTHROPIC_API_KEY=your_api_key_here
```

Keep the API key private. The contract workflow sends up to the first 3,000
characters of extracted contract text to Anthropic.

## Usage

Run commands from the project folder.

### Summarize the sample contract and generate the sales report

```powershell
python Documents.py
```

This is the default `documents` workflow. It reads `supplier_contract.pdf`,
prints the contract summary, and saves `november_report.docx`.

### Index a folder of PDFs

```powershell
python Documents.py batch invoices
```

This scans PDF files directly inside the `invoices` folder and writes
`invoice_index.csv` in the project folder. Specify a different input folder or
output path if needed:

```powershell
python Documents.py batch "C:\path\to\invoices" --output "C:\path\to\invoice_index.csv"
```

The CSV columns are `file`, `pages`, `chars`, and `preview`. Files are
processed in filename order, and the preview contains up to 100 characters of
extracted text. Batch indexing does not make an Anthropic API call or require
an API key.

## Files

- `Documents.py` — workflows and command-line entry point.
- `supplier_contract.pdf` — fictional sample for text-extraction testing; it is
  not a legal contract.
- `november_report.docx` — generated sales report (overwritten when the
  default workflow completes).
- `invoices/` — create this folder and add PDFs before running batch indexing.

## Notes

- `pypdf` extracts embedded text; it does not perform OCR. Image-only scanned
  PDFs need OCR before their text can be extracted.
- The sales figures in the Word report are hard-coded demonstration data.
- The default workflow needs a readable `supplier_contract.pdf` and a valid
  `ANTHROPIC_API_KEY`. The batch workflow reports an error if its input folder
  does not exist.
