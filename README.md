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

## Sample Results

| Metric | Result |
|---|---:|
| Total Orders | 10 |
| Unique Products | 4 |
| Units Sold | 24 |
| Total Sales | 805.50 |

## Excel Workbook Sheets

- **Overview:** Key sales metrics
- **Sales Details:** Individual order records
- **Product Summary:** Units sold and revenue by product

## Download Sample Excel Report

Download and explore the generated Excel sales report:

[Download Sample Sales Report](sample_sales_report.xlsx)

The workbook includes:
- Overview of sales metrics
- Detailed sales records
- Product-wise sales summary
