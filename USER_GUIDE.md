# Speedy Invoice Extractor User Guide

## Introduction
The Speedy Invoice Extractor is a desktop application designed to automate the extraction of invoice data from Speedy CFS PDF invoices and correlate it with the Export Job Register to generate a ready-to-upload Logisys Purchase CSV file.

## How to Use

### 1. Launching the App
Double click the `Speedy_Invoice_Extractor.exe` application to start the program. It will open in full window mode with the Nagarkot Forwarders branding.

### 2. The Workflow (Step-by-Step)
1. **Select Export Job Register**: Click the button to select your master job register.
   - *Note: The Excel file must have a `.xlsx` extension and contain column headers like "SB No." or "Shipping Bill" and "Job No" for successful lookup.*
2. **Select PDF Invoices**: Click the button to locate your Speedy invoices.
   - *Note: Only valid `.pdf` files can be selected.*
3. **Process & Generate CSV**: Click the "▶ Process & Generate CSV" button to start extraction and data merging.
   - *Note: Your final output file (`Speedy_Purchases_DDMMYYYY_HHMM.csv`) will be automatically saved inside a new `CSV Output` folder located in the same directory as your selected PDFs.*

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| Select Export Job Register | Button to choose the master Excel file for Job Number lookup. | `.xlsx` file |
| Select PDF Invoices | Button to choose multiple Speedy invoice files. | `.pdf` files |
| ▶ Process & Generate CSV | Starts the PDF parsing, job number lookup, and CSV creation. | N/A |
| Processing Log | Text area displaying real-time extraction details and errors for each file. | Read-Only |

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| Please select the Export Job Register first. | You attempted to process before attaching the `.xlsx` register file. | Click "Select Export Job Register" and pick a file. |
| Please select at least one PDF invoice. | You attempted to process without providing any invoices. | Click "Select PDF Invoices" and select valid `.pdf` files. |
| Failed to load: [Error message] | The application could not read the selected `.xlsx` file. | Check if the Excel is corrupted, open in another program, or not in `.xlsx` format. |
| Invoice No not found in PDF. | The parser couldn't identify the "Invoice No :" in the document. | Ensure the uploaded PDF is a valid Speedy Invoice, not a generalized receipt. |
| Error reading PDF: [Error message] | The PDF structure is broken or couldn't be evaluated. | Ensure the file isn't password protected or an unscrapable image. |
