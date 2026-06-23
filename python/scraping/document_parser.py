"""
Document Parser - HTML -> Structured JSON (Portfolio Example)
=============================================================
This script reads a web page using CSS selectors and turns it into clean,
structured JSON. The idea:

    "Point CSS selectors at the page's headings and body, and the document
     comes out as tidy JSON (title + sections)."

This is the core pattern behind 'HTML parser' style work: you write ONE parser
per source (site), mostly by changing the selectors below. ~90% of the work is
the SELECTORS dict at the top of this file.

Usage:
    pip install requests beautifulsoup4
    python document_parser.py                 # parses the example URL (online)
    python document_parser.py page.html       # parses a saved local HTML file (offline)
"""

import sys
import json

import requests
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------------
# 1) SELECTOR CONFIGURATION
#    This is the heart of the work: we define the "address" (CSS selector) of
#    each part. For a new source you usually change ONLY these three lines.
# ---------------------------------------------------------------------------
SELECTORS = {
    "title":       "div.product_main h1",        # document title (single element)
    "description": "#product_description + p",    # description: the <p> right after #product_description
    "info_rows":   "table.table-striped tr",      # rows of the "Product Information" table (th + td)
}

# Default example page used when no source is passed on the command line:
DEFAULT_URL = "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"


def load_soup(source: str) -> BeautifulSoup:
    """Fetch the source if it's a URL, otherwise read it from disk; return BeautifulSoup."""
    if source.startswith("http"):
        # --- Online: download the page ---
        resp = requests.get(source, timeout=15)
        resp.raise_for_status()
        resp.encoding = "utf-8"   # pin encoding to UTF-8 to avoid mojibake like "Â£"
        html = resp.text
    else:
        # --- Offline: read a saved HTML file (this mirrors the real workflow) ---
        with open(source, encoding="utf-8") as f:
            html = f.read()
    return BeautifulSoup(html, "html.parser")


def parse_document(soup: BeautifulSoup) -> dict:
    """Convert the page into a structured dict (JSON) based on SELECTORS."""

    # --- Title: a single element -> .select_one ---
    title_el = soup.select_one(SELECTORS["title"])
    title = title_el.get_text(strip=True) if title_el else None

    sections = []

    # --- Section 1: Description ---
    desc_el = soup.select_one(SELECTORS["description"])
    if desc_el:
        sections.append({
            "heading": "Description",
            "content": [desc_el.get_text(strip=True)],
        })

    # --- Section 2: Product Information table (each row: label -> value) ---
    info_lines = []
    for row in soup.select(SELECTORS["info_rows"]):     # .select -> returns multiple elements
        label_el = row.select_one("th")
        value_el = row.select_one("td")
        if label_el and value_el:
            label = label_el.get_text(strip=True)
            value = value_el.get_text(strip=True)
            info_lines.append(f"{label}: {value}")
    if info_lines:
        sections.append({
            "heading": "Product Information",
            "content": info_lines,
        })

    # --- Assemble everything into a clean, structured JSON shape ---
    return {
        "title": title,
        "section_count": len(sections),
        "sections": sections,
    }


def main():
    # Use the default URL when no source is given on the command line
    source = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_URL
    print(f"[*] Parsing: {source}")

    soup = load_soup(source)
    document = parse_document(soup)

    # Convert to JSON. ensure_ascii=False keeps non-ASCII characters intact
    output = json.dumps(document, indent=2, ensure_ascii=False)

    # Print to screen and save to a file
    print(output)
    with open("document.json", "w", encoding="utf-8") as f:
        f.write(output)
    print("\n[+] Wrote document.json.")


if __name__ == "__main__":
    main()
