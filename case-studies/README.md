# Case Studies

Real-world data analysis projects demonstrating Excel, Python, and SQL skills.

## 1. Turkey Inflation & Unemployment Dashboard (2019–2025)

**File:** `turkey_macro_dashboard.xlsx`
**Generator:** `turkey_macro_dashboard.py`

A macroeconomic dashboard analyzing Turkey's CPI inflation and unemployment trends over 76 months.

### Sheets

| Sheet | Content |
|-------|---------|
| **Ham Veri** | 76 rows of monthly data (Jan 2019 – Apr 2025) with conditional color scales |
| **Analiz** | Yearly averages, min/max, CORREL() analysis, summary statistics |
| **Dashboard** | KPI cards, dual line charts, yearly comparison table, correlation analysis |

### Features Demonstrated

- Advanced Excel formulas: `CORREL()`, `AVERAGEIFS`
- Conditional formatting (3-color scales)
- Professional chart design (line charts with custom styling)
- KPI dashboard layout
- Multi-sheet workbook architecture
- Real-world data from TÜİK & TCMB

### Key Finding

Negative correlation between inflation and unemployment — a Phillips Curve pattern amplified by Turkey's currency crisis. Peak inflation reached 85.51% (Oct 2022) while unemployment dropped to single digits.

### Regenerate

```bash
pip install openpyxl
python turkey_macro_dashboard.py
```

---

## 2. E-Commerce Sales Analysis (2023-2024)

**File:** `ecommerce_analysis.xlsx`
**Generator:** `ecommerce_analysis.py`

A comprehensive e-commerce sales analysis covering 550+ transactions across 6 international markets over 2 years, with RFM customer segmentation, product performance metrics, and an executive dashboard.

### Sheets

| Sheet | Content |
|-------|---------|
| **Raw Data** | 550 rows of transaction data with formulas for revenue calculation, auto-filters, freeze panes |
| **Product Performance** | SUMIFS/COUNTIFS-based product ranking with conditional formatting (green=top 5, red=bottom 5) |
| **Customer Segmentation** | RFM-style analysis with 4 segments: VIP, Regular, At Risk, New — color-coded rows |
| **Monthly Trends** | 24-month summary with line chart (revenue) and bar chart (order count) |
| **Dashboard** | Executive view with KPI cards, country/category breakdowns, and key insights |
| **Country Analysis** | Revenue distribution by country with bar chart, pie chart, and top 3 products per market |

### Features Demonstrated

- Advanced Excel formulas: `SUMPRODUCT`, `COUNTIF`, `SUMIFS` with cross-sheet references
- RFM-style customer segmentation (Recency, Frequency, Monetary)
- Multiple chart types: line, bar, pie with professional styling
- Conditional formatting and color-coded segments
- KPI dashboard layout with blue/gray professional color scheme
- Multi-market analysis (USA, UK, Germany, France, Turkey, Canada)
- Reproducible data generation with `random.seed(42)`
- 6-sheet workbook architecture

### Key Findings

- 550 transactions from 150 unique customers across 6 countries generated significant revenue over the 2-year period.
- USA leads all markets in revenue share, followed by UK and Turkey.
- Electronics and Clothing are the top-performing product categories.
- Credit Card is the dominant payment method across all markets.
- VIP customers (top 10% by spend) drive a disproportionate share of total revenue.

### Regenerate

```bash
pip install openpyxl
python ecommerce_analysis.py
```

---

## 3. Financial Budget Tracker (2024)

**File:** `budget_tracker.xlsx`
**Generator:** `budget_tracker.py`

A comprehensive personal/small business budget tracker covering 12 months (Jan-Dec 2024) with 382 realistic income and expense transactions, budget vs actual analysis, and an executive dashboard.

### Sheets

| Sheet | Content |
|-------|---------|
| **Transactions** | 382 rows of income/expense transactions with date, category, subcategory, description, amount, and payment method. Auto-filters, freeze panes, alternating row colors |
| **Monthly Budget** | Budget vs Actual comparison for 8 expense categories across all 12 months with variance and variance % columns. Conditional formatting (green = under budget, red = over budget) |
| **Monthly Summary** | Monthly totals for income, expenses, net savings, and savings rate. Line chart (income vs expenses) and bar chart (monthly net savings) |
| **Category Breakdown** | Each expense category ranked by total spent with % of total, monthly average, annual budget, and over/under status. Pie chart and horizontal bar chart (budget vs actual) |
| **Dashboard** | Executive view with 6 KPI cards (Total Income, Total Expenses, Net Savings, Savings Rate, Best Month, Most Expensive Category), 3 charts, top 5 categories table, budget health indicator, and key insights |
| **Yearly Report** | Annual summary, quarterly breakdown (Q1-Q4), year-end projections, category budget comparison with status, and actionable recommendations |

### Features Demonstrated

- Advanced Excel formulas and cross-sheet data architecture
- Budget vs Actual variance analysis with conditional formatting
- Multiple chart types: line, bar (column & horizontal), pie with custom styling
- KPI dashboard layout with professional color scheme (dark blue/green/red)
- Conditional formatting for over/under budget indicators
- Realistic transaction generation with seasonal spending patterns
- Currency ($#,##0.00) and percentage formatting throughout
- 6-sheet workbook with 382 reproducible transactions (`random.seed(42)`)

### Key Findings

- 382 transactions generated across 12 months covering 8 expense categories and 4 income sources.
- Housing is consistently the largest expense category, driven by recurring rent and utilities.
- Holiday season (Nov-Dec) shows increased shopping spend due to seasonal patterns.
- Salary provides the stable income base, supplemented by variable freelance earnings and quarterly investment dividends.
- The budget health indicator and category-level variance analysis highlight areas needing attention.

### Regenerate

```bash
pip install openpyxl
python budget_tracker.py
```
