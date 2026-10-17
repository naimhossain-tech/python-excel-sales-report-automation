import unittest
from pathlib import Path
from decimal import Decimal

from report_generator import load_sales

class TestSalesReport(unittest.TestCase):
    def setUp(self):
        self.csv_path = Path("data/sales_january.csv")

    def test_loads_all_sales_records(self):
        records = load_sales(self.csv_path)
        self.assertEqual(len(records), 10)

    def test_calculates_total_sales(self):
        records = load_sales(self.csv_path)
        total = sum(item["total"] for item in records)
        self.assertEqual(total, Decimal("805.50"))

    def test_calculates_total_units(self):
        records = load_sales(self.csv_path)
        units = sum(item["quantity"] for item in records)
        self.assertEqual(units, 24)

if __name__ == "__main__":
    unittest.main()
