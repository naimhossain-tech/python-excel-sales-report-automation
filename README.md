# Python Excel Sales Report Automation

A Python automation project that converts CSV sales data into a formatted Excel report.

## Features
- Reads and validates CSV sales data
- Calculates total orders, units sold, and sales revenue
- Creates an Excel workbook with Overview, Sales Details, and Product Summary sheets
- Formats headers, currency values, and filters
- Includes automated unit tests

## Requirements
- Python
- openpyxl 3.1.5

## Run the Project
```bash
pip install -r requirements.txt
python report_generator.py
```

## Run Tests
```bash
python -m unittest discover -s tests -v
```

## Sample Data
Sample CSV data is available in data/sales_january.csv.

## Disclaimer
This project uses sample sales data for demonstration purposes.
