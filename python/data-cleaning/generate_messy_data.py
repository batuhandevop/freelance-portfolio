"""
generate_messy_data.py
---------------------
Generates a realistic messy sales CSV file for demonstrating data cleaning.

The output intentionally contains every common data-quality problem you would
encounter in a real client dataset:
  - duplicate rows
  - missing values (empty, "N/A", "null", "none")
  - dates in multiple formats
  - extra whitespace
  - mixed-case / inconsistent text
  - typos in category names
  - outlier and nonsensical values
  - numbers stored as strings (with $, commas)
"""

import csv
import random
import os

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "messy_sales_data.csv")
NUM_ROWS = 200
SEED = 42

random.seed(SEED)

# ---------------------------------------------------------------------------
# Reference pools
# ---------------------------------------------------------------------------
CLEAN_CATEGORIES = ["Electronics", "Clothing", "Home & Garden",
                    "Sports", "Books", "Toys"]

# Deliberately messy variants (typos, casing, whitespace)
CATEGORY_VARIANTS = {
    "Electronics":  ["Electronics", "electronics", "ELECTRONICS",
                     " Electronics ", "Electornics", "Elecronics"],
    "Clothing":     ["Clothing", "clothing", "CLOTHING",
                     "Clothng", "  Clothing"],
    "Home & Garden": ["Home & Garden", "home & garden", "Home &  Garden",
                      "Home and Garden", "Home&Garden"],
    "Sports":       ["Sports", "sports", "SPORTS", "Sprots", " Sports "],
    "Books":        ["Books", "books", "BOOKS", "Boks"],
    "Toys":         ["Toys", "toys", "TOYS", "Tys"],
}

REGIONS = ["North", "South", "East", "West"]
REGION_VARIANTS = {
    "North": ["North", "north", " North", "NORTH", "N"],
    "South": ["South", "south", "  South ", "SOUTH", "S"],
    "East":  ["East", "east", "EAST", " East"],
    "West":  ["West", "west", "WEST", "West "],
}

SALESPERSON_NAMES = [
    "Alice Johnson", "Bob Smith", "Carol White", "David Brown",
    "Eva Martinez", "Frank Lee", "Grace Kim", "Hank Patel",
]

MISSING_MARKERS = ["", "N/A", "null", "none", "NA", "-"]


def random_date():
    """Return a date string in one of several inconsistent formats."""
    year = random.choice([2023, 2024, 2025])
    month = random.randint(1, 12)
    day = random.randint(1, 28)

    fmt = random.choice(["iso", "us", "eu", "mon"])
    if fmt == "iso":
        return f"{year}-{month:02d}-{day:02d}"
    elif fmt == "us":
        return f"{month:02d}/{day:02d}/{year}"
    elif fmt == "eu":
        return f"{day:02d}.{month:02d}.{year}"
    else:
        import datetime
        d = datetime.date(year, month, day)
        return d.strftime("%d-%b-%Y")


def random_price():
    """Return a price that is sometimes formatted oddly."""
    base = round(random.uniform(5, 500), 2)
    style = random.choice(["clean", "dollar", "comma", "negative", "string"])
    if style == "clean":
        return str(base)
    elif style == "dollar":
        return f"${base:,.2f}"
    elif style == "comma":
        return f"{base:,.2f}"
    elif style == "negative":
        # Outlier: negative price
        return str(-abs(base))
    else:
        return f" {base} "


def random_quantity():
    """Return a quantity, sometimes nonsensical."""
    roll = random.random()
    if roll < 0.05:
        return str(random.randint(-10, -1))   # negative quantity
    elif roll < 0.10:
        return str(random.randint(5000, 99999))  # unrealistic outlier
    elif roll < 0.15:
        return str(round(random.uniform(1, 10), 2))  # float instead of int
    else:
        return str(random.randint(1, 100))


def maybe_missing(value, probability=0.08):
    """With some probability, replace a value with a missing-data marker."""
    if random.random() < probability:
        return random.choice(MISSING_MARKERS)
    return value


def generate_row(order_id):
    """Generate a single messy row."""
    cat_key = random.choice(CLEAN_CATEGORIES)
    region_key = random.choice(REGIONS)

    return {
        "order_id":     str(order_id),
        "date":         maybe_missing(random_date(), 0.06),
        "category":     maybe_missing(random.choice(CATEGORY_VARIANTS[cat_key]), 0.05),
        "product_name": maybe_missing(
            random.choice([
                "Widget A", "widget a", "Widget  A",
                "Gadget Pro", "gadget pro", " Gadget Pro",
                "Super Deluxe Set", "super deluxe set",
                "Basic Kit", "BASIC KIT", "basic kit ",
                "Premium Bundle", "  Premium Bundle  ",
            ]),
            0.05,
        ),
        "quantity":     maybe_missing(random_quantity(), 0.07),
        "unit_price":   maybe_missing(random_price(), 0.07),
        "region":       maybe_missing(random.choice(REGION_VARIANTS[region_key]), 0.05),
        "salesperson":  maybe_missing(random.choice(SALESPERSON_NAMES), 0.06),
    }


def main():
    fieldnames = ["order_id", "date", "category", "product_name",
                  "quantity", "unit_price", "region", "salesperson"]

    rows = [generate_row(i + 1001) for i in range(NUM_ROWS)]

    # Inject exact duplicate rows (pick ~5 % of rows and duplicate them)
    duplicates = random.sample(rows, k=int(NUM_ROWS * 0.05))
    rows.extend(duplicates)
    random.shuffle(rows)

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Generated {len(rows)} rows (including duplicates) -> {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
