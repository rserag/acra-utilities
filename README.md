# ACRA loan Data to Excel Converter

This Python application converts loan data in PDF format into an Excel file (`.xlsx`), with the following features:

- Parses loan details such as loan number, loan conditions, and amounts.
- Extracts and formats data fields, including currency values (in Armenian Dram).
- Automatically calculates totals for numeric fields.
- Outputs the results to a well-organized Excel sheet.

## Prerequisites

- Python 3.6+ 
- Required Python libraries:
  - `openpyxl`: For creating and formatting the Excel file.
  - `pypdf`: For extracting text from PDF files.
