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
 
### Step 1: Install Python

Ensure you have Python 3.8 or higher installed. You can download it from [python.org](https://www.python.org/downloads/).

### Step 2: Install Dependencies

The repository contains a `requirements.txt` file that lists all the required libraries. To install the dependencies:

1. Open a terminal or command prompt.
2. Navigate to the directory containing the `requirements.txt` file.
3. Run the following command:

   ```bash
   pip install -r requirements.txt
   ```

## Usage
- Run the script from the command line:

```bash
python total_received_loans.py -n <PDF_FILE_NAME> [--debug]
```

### Arguments
- -n, --name (required): The name of the PDF file to process.
- --debug (optional): Enable debug logging to print additional information during execution.

## Example
```bash
python total_received_loans.py -n loans.pdf --debug
```
This command processes a PDF file named loans.pdf and creates an Excel file named loans.xlsx.

## Output
- The script generates an Excel file named <PDF_FILE_NAME>.xlsx (e.g., loans.xlsx).
- The file contains:
   - Loan numbers in the first column.
   - Extracted details from the PDF in subsequent columns.
   - A "Totals" row for numeric columns at the bottom.

## Notes
- The script uses a regular expression to extract blocks of text between the terms N ՎԱՐԿ and Նշումներ գրավի վերաբերյալ.
- Ensure the PDF contains properly formatted data for accurate processing.

## Debugging
- Use the --debug flag to enable debug messages, which help in troubleshooting issues with data extraction or formatting.

## License
This project is open-source and free to use under the MIT License.
