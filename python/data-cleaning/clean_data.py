"""
clean_data.py
-------------
Reads messy_sales_data.csv, applies a full cleaning pipeline, and writes
cleaned_sales_data.csv.  Prints a before/after summary report to the console.

Every step is commented so the reader can follow the reasoning.
Only pandas + the standard library are used.
"""

import os
import re
from datetime import datetime

import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(BASE_DIR, "messy_sales_data.csv")
OUTPUT_FILE = os.path.join(BASE_DIR, "cleaned_sales_data.csv")


# ===================================================================
# STEP 0 : Load the raw data
# ===================================================================
def load_data(path: str) -> pd.DataFrame:
    """Read CSV and treat common missing-value markers as NaN right away."""
    return pd.read_csv(
        path,
        na_values=["N/A", "null", "none", "NA", "-", ""],
        keep_default_na=True,
    )


# ===================================================================
# STEP 1 : Remove duplicate rows
# ===================================================================
def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Drop exact duplicate rows, keeping the first occurrence."""
    before = len(df)
    df = df.drop_duplicates()
    after = len(df)
    print(f"  [Duplicates]   Removed {before - after} duplicate rows.")
    return df.reset_index(drop=True)


# ===================================================================
# STEP 2 : Strip whitespace from all string columns
# ===================================================================
def strip_whitespace(df: pd.DataFrame) -> pd.DataFrame:
    """Remove leading/trailing spaces and collapse internal runs of spaces."""
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = (
            df[col]
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )
    print(f"  [Whitespace]   Stripped whitespace in {len(str_cols)} text columns.")
    return df


# ===================================================================
# STEP 3 : Standardize date column
# ===================================================================
_DATE_PATTERNS = [
    # ISO:  2024-03-15
    (r"^\d{4}-\d{2}-\d{2}$", "%Y-%m-%d"),
    # US:   03/15/2024
    (r"^\d{2}/\d{2}/\d{4}$", "%m/%d/%Y"),
    # EU:   15.03.2024
    (r"^\d{2}\.\d{2}\.\d{4}$", "%d.%m.%Y"),
    # Mon:  15-Mar-2024
    (r"^\d{2}-[A-Za-z]{3}-\d{4}$", "%d-%b-%Y"),
]


def _parse_date(value):
    """Try each known date pattern and return a datetime or NaT."""
    if pd.isna(value):
        return pd.NaT
    value = str(value).strip()
    for regex, fmt in _DATE_PATTERNS:
        if re.match(regex, value):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
    return pd.NaT


def standardize_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Parse dates from multiple formats into a single YYYY-MM-DD format."""
    unparsed_before = df["date"].isna().sum()
    df["date"] = df["date"].apply(_parse_date)
    unparsed_after = df["date"].isna().sum()
    new_failures = unparsed_after - unparsed_before
    if new_failures > 0:
        print(f"  [Dates]        Standardized dates; {new_failures} could not be parsed.")
    else:
        print(f"  [Dates]        All non-null dates parsed successfully.")
    return df


# ===================================================================
# STEP 4 : Clean and standardize category names
# ===================================================================
# Mapping of known typos / variants to their canonical form.
CATEGORY_MAP = {
    "electronics":   "Electronics",
    "electornics":   "Electronics",
    "elecronics":    "Electronics",
    "clothing":      "Clothing",
    "clothng":       "Clothing",
    "home & garden": "Home & Garden",
    "home and garden": "Home & Garden",
    "home&garden":   "Home & Garden",
    "sports":        "Sports",
    "sprots":        "Sports",
    "books":         "Books",
    "boks":          "Books",
    "toys":          "Toys",
    "tys":           "Toys",
}


def clean_categories(df: pd.DataFrame) -> pd.DataFrame:
    """Fix typos and inconsistent casing in the category column."""
    fixed = 0
    def _map(val):
        nonlocal fixed
        if pd.isna(val):
            return val
        lookup = val.strip().lower()
        canonical = CATEGORY_MAP.get(lookup)
        if canonical and canonical != val:
            fixed += 1
        return canonical if canonical else val.strip().title()

    df["category"] = df["category"].apply(_map)
    print(f"  [Categories]   Fixed {fixed} inconsistent category names.")
    return df


# ===================================================================
# STEP 5 : Normalize product names
# ===================================================================
def clean_product_names(df: pd.DataFrame) -> pd.DataFrame:
    """Title-case product names after stripping whitespace (already done)."""
    df["product_name"] = df["product_name"].str.title()
    print("  [Products]     Standardized product name casing.")
    return df


# ===================================================================
# STEP 6 : Normalize region names
# ===================================================================
REGION_MAP = {
    "n": "North", "north": "North",
    "s": "South", "south": "South",
    "e": "East",  "east":  "East",
    "w": "West",  "west":  "West",
}


def clean_regions(df: pd.DataFrame) -> pd.DataFrame:
    """Map abbreviations and fix casing for region values."""
    def _map(val):
        if pd.isna(val):
            return val
        return REGION_MAP.get(val.strip().lower(), val.strip().title())

    df["region"] = df["region"].apply(_map)
    print("  [Regions]      Standardized region names.")
    return df


# ===================================================================
# STEP 7 : Clean and convert unit_price to float
# ===================================================================
def clean_unit_price(df: pd.DataFrame) -> pd.DataFrame:
    """Remove currency symbols and commas, then cast to float.
    Negative prices are replaced with NaN (they are data-entry errors)."""
    def _parse_price(val):
        if pd.isna(val):
            return float("nan")
        s = str(val).strip().replace("$", "").replace(",", "")
        try:
            price = float(s)
            return price if price >= 0 else float("nan")
        except ValueError:
            return float("nan")

    negatives = df["unit_price"].astype(str).str.startswith("-").sum()
    df["unit_price"] = df["unit_price"].apply(_parse_price)
    print(f"  [Prices]       Converted to float; replaced {negatives} negative prices with NaN.")
    return df


# ===================================================================
# STEP 8 : Clean and convert quantity to int
# ===================================================================
def clean_quantity(df: pd.DataFrame) -> pd.DataFrame:
    """Cast quantity to numeric. Negatives and extreme outliers (>1000)
    are treated as data-entry errors and set to NaN."""
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")

    outlier_mask = (df["quantity"] < 0) | (df["quantity"] > 1000)
    outlier_count = outlier_mask.sum()
    df.loc[outlier_mask, "quantity"] = float("nan")

    # Convert remaining floats to nullable integer (keeps NaN support)
    df["quantity"] = df["quantity"].round(0).astype("Int64")
    print(f"  [Quantities]   Converted to int; flagged {outlier_count} outliers as NaN.")
    return df


# ===================================================================
# STEP 9 : Handle remaining missing values
# ===================================================================
def handle_missing(df: pd.DataFrame) -> pd.DataFrame:
    """Report on remaining NaN counts per column.
    We keep NaN rows rather than dropping them -- in real projects the
    strategy (drop, fill, impute) depends on the business context."""
    missing = df.isna().sum()
    total = missing.sum()
    if total:
        print(f"  [Missing]      {total} NaN values remain across columns:")
        for col, n in missing.items():
            if n > 0:
                print(f"                   {col}: {n}")
    else:
        print("  [Missing]      No missing values remain.")
    return df


# ===================================================================
# Summary report
# ===================================================================
def print_report(raw: pd.DataFrame, clean: pd.DataFrame):
    """Print a side-by-side before/after summary."""
    sep = "=" * 60
    print(f"\n{sep}")
    print("  BEFORE / AFTER SUMMARY")
    print(sep)
    print(f"  {'Metric':<30} {'Before':>10} {'After':>10}")
    print(f"  {'-'*30} {'-'*10} {'-'*10}")
    print(f"  {'Total rows':<30} {len(raw):>10} {len(clean):>10}")
    print(f"  {'Duplicate rows':<30} {len(raw) - len(raw.drop_duplicates()):>10} {'0':>10}")
    print(f"  {'Total NaN cells':<30} {raw.isna().sum().sum():>10} {clean.isna().sum().sum():>10}")

    # Per-column null counts
    print(f"\n  Null counts per column:")
    print(f"  {'Column':<20} {'Before':>10} {'After':>10}")
    print(f"  {'-'*20} {'-'*10} {'-'*10}")
    for col in clean.columns:
        b = raw[col].isna().sum() if col in raw.columns else "n/a"
        a = clean[col].isna().sum()
        print(f"  {col:<20} {b:>10} {a:>10}")

    # Data types
    print(f"\n  Data types after cleaning:")
    for col in clean.columns:
        print(f"    {col:<20} {str(clean[col].dtype)}")

    # Quick value-counts for categorical columns
    for col in ["category", "region"]:
        print(f"\n  Unique values in '{col}':")
        for val, cnt in clean[col].value_counts(dropna=False).items():
            label = val if pd.notna(val) else "<missing>"
            print(f"    {label:<25} {cnt}")

    print(f"\n{sep}\n")


# ===================================================================
# Main pipeline
# ===================================================================
def main():
    print("Loading data ...\n")
    raw = load_data(INPUT_FILE)

    # Keep an unmodified copy for the report
    raw_snapshot = raw.copy()

    print("Cleaning pipeline:")
    df = raw.copy()
    df = remove_duplicates(df)
    df = strip_whitespace(df)
    df = standardize_dates(df)
    df = clean_categories(df)
    df = clean_product_names(df)
    df = clean_regions(df)
    df = clean_unit_price(df)
    df = clean_quantity(df)
    df = handle_missing(df)

    # Write the clean file
    df.to_csv(OUTPUT_FILE, index=False, date_format="%Y-%m-%d")
    print(f"\nCleaned data written to {OUTPUT_FILE}")

    # Summary
    print_report(raw_snapshot, df)


if __name__ == "__main__":
    main()
