# Book Scraper - Web Scraping Portfolio Example

A Python web scraper that collects book data from [books.toscrape.com](https://books.toscrape.com), a sandbox site designed specifically for practicing web scraping techniques.

## What It Does

- Scrapes book listings across multiple pages (title, price, rating, availability, category)
- Handles pagination automatically
- Exports results to both CSV and Excel formats
- Includes retry logic for unreliable network conditions

## How to Run

```bash
pip install requests beautifulsoup4 openpyxl
python book_scraper.py
```

Output files are saved to the same directory: `books.csv` and `books.xlsx`.

## A Note on Ethical Scraping

This project targets a site that explicitly permits scraping. When scraping real-world sites, always:

- Check `robots.txt` and the site's Terms of Service before scraping
- Use polite request intervals to avoid overloading servers
- Identify your scraper with a descriptive `User-Agent` header
- Cache responses when possible to minimize repeat requests
- Never scrape personal or sensitive data without consent

## Tech Stack

- **requests** - HTTP client
- **BeautifulSoup4** - HTML parsing
- **openpyxl** - Excel file generation
