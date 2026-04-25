# Excel Formula Toolkit

A Python-based generator that creates `formula_toolkit.xlsx` -- an Excel workbook
demonstrating advanced formula patterns commonly needed in business analytics and
data processing work.

## Sheets

| Sheet | Focus | Key Formulas |
|-------|-------|--------------|
| LOOKUP Examples | Data retrieval from tables | VLOOKUP, INDEX-MATCH, XLOOKUP |
| Conditional Formulas | Aggregation with criteria | SUMIFS, COUNTIFS, AVERAGEIFS |
| Text Functions | Messy data cleanup | CONCATENATE, LEFT, RIGHT, MID, TRIM, SUBSTITUTE |
| Date Functions | Project timeline math | NETWORKDAYS, EOMONTH, DATEDIF, WORKDAY |

## Usage

```bash
pip install openpyxl
python formula_toolkit.py
```

This generates `formula_toolkit.xlsx` in the current directory.

## Requirements

- Python 3.8+
- openpyxl
