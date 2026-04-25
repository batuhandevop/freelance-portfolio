# Sales Dashboard - Excel Portfolio Piece

An automated Excel sales dashboard generated via Python (openpyxl). This project demonstrates the ability to build professional, formula-driven dashboards entirely through code -- useful for recurring reports, client deliverables, and data pipelines that end in Excel.

## What the Dashboard Shows

**Data Sheet** -- 120 rows of realistic 2024 sales transactions across four regions, five product categories, and six salesperson names.

**Dashboard Sheet:**

- **KPI Summary** -- Total Revenue, Total Orders, and Average Order Value calculated with Excel formulas (SUM, COUNTA, AVERAGE).
- **Revenue by Region** -- SUMIFS-based breakdown for North, South, East, and West regions.
- **Revenue by Product Category** -- SUMIFS-based breakdown for each product line (Laptops, Peripherals, Accessories, etc.).
- **Monthly Revenue Trend** -- SUMIFS pulling monthly totals across all 12 months of 2024.
- **Conditional Formatting** -- Color scales on revenue columns and data bars on quantity columns for at-a-glance analysis.
- **Professional Styling** -- Header colors, borders, number formatting, and column widths tuned for readability.

## How to Generate

```bash
pip install openpyxl
python sales_dashboard.py
```

This produces `sales_dashboard.xlsx` in the current directory. Open it in Excel, LibreOffice Calc, or Google Sheets.

## Screenshot Placeholder

> _Insert screenshot of the finished Dashboard sheet here._

## Skills Demonstrated

- Python automation of Excel workbooks
- Excel formula construction (SUM, SUMIFS, COUNTA, AVERAGE)
- Conditional formatting via openpyxl rules
- Clean data modeling and realistic sample data generation
