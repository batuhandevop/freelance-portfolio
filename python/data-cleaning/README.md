# Data Cleaning Pipeline - Sales Data

A demonstration of practical data cleaning techniques using Python and pandas.

## Files

| File | Purpose |
|---|---|
| `generate_messy_data.py` | Generates a realistic messy CSV (`messy_sales_data.csv`) containing common real-world data problems |
| `clean_data.py` | Reads the messy CSV, fixes every issue, and outputs `cleaned_sales_data.csv` with a before/after report |

## Data Problems Handled

- Duplicate rows
- Missing values in various forms (`N/A`, `null`, empty strings, `NaN`)
- Dates in inconsistent formats (`YYYY-MM-DD`, `MM/DD/YYYY`, `DD-Mon-YYYY`)
- Extra leading/trailing whitespace
- Mixed-case text and inconsistent capitalization
- Typos and misspellings in category names
- Outlier values (negative prices, unrealistic quantities)
- Numbers stored as strings with embedded characters (`$`, commas)

## Usage

```bash
python generate_messy_data.py   # creates messy_sales_data.csv
python clean_data.py            # cleans it and creates cleaned_sales_data.csv
```

## Requirements

- Python 3.8+
- pandas
