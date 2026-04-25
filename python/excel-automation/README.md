# Excel Report Automation

Automated monthly sales report generator built with Python and openpyxl.

## What It Does

Takes raw monthly sales data (by region and product) and produces a polished,
multi-sheet Excel workbook (`monthly_report.xlsx`) containing:

- **Cover Page** -- report title, company name, generation date
- **Summary** -- KPI dashboard with total revenue, growth %, top product, top region (formula-driven)
- **Detail** -- full data table with filters, freeze panes, and conditional formatting
- **Per-Region Sheets** -- auto-generated breakdown for each region
- **Charts** -- bar chart (revenue by product) and line chart (monthly trend)

## Usage

```bash
pip install openpyxl
python report_generator.py
```

Output: `monthly_report.xlsx` in the same directory.

## Dependencies

- Python 3.8+
- openpyxl
