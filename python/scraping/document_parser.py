"""
Document Parser - HTML -> Structured JSON (Portfolio Example)
=============================================================
Bu script bir web sayfasini CSS selector'lar kullanarak okur ve temiz,
yapilandirilmis JSON'a cevirir. Mantik su:

    "Sayfadaki basliklara ve govdeye CSS selector'larla isaret et,
     belge derli toplu JSON (baslik + bolumler) olarak ciksin."

Bu, 'HTML parser' tarzi islerin cekirdek desenidir: her kaynak (site) icin
TEK bir parser yazarsin -- cogunlukla sadece asagidaki selector'lari
degistirerek. Isin %90'i bu dosyanin en ustundeki SELECTORS sozlugudur.

Calistirma:
    pip install requests beautifulsoup4
    python document_parser.py                 # ornek URL'i parse eder (online)
    python document_parser.py page.html       # kaydedilmis yerel HTML'i parse eder (offline)
"""

import sys
import json

import requests
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------------
# 1) SELECTOR YAPILANDIRMASI
#    Isin kalbi burasi: her parcanin "adresini" (CSS selector) tanimliyoruz.
#    Yeni bir kaynak icin genelde SADECE bu uc satiri degistirirsin.
# ---------------------------------------------------------------------------
SELECTORS = {
    "title":       "div.product_main h1",        # belge basligi (tek eleman)
    "description": "#product_description + p",    # aciklama: #product_description'dan HEMEN sonraki <p>
    "info_rows":   "table.table-striped tr",      # "Product Information" tablosunun satirlari (th + td)
}

# Komut satirindan kaynak verilmezse kullanilacak varsayilan ornek sayfa:
DEFAULT_URL = "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"


def load_soup(source: str) -> BeautifulSoup:
    """Kaynak URL ise indir, yerel dosya ise diskten oku; BeautifulSoup dondur."""
    if source.startswith("http"):
        # --- Online: sayfayi indir ---
        resp = requests.get(source, timeout=15)
        resp.raise_for_status()
        resp.encoding = "utf-8"   # encoding'i acikca UTF-8'e sabitle -> "Â£" gibi bozulmalari onler
        html = resp.text
    else:
        # --- Offline: kaydedilmis HTML dosyasini oku (gercek isin akisi boyle) ---
        with open(source, encoding="utf-8") as f:
            html = f.read()
    return BeautifulSoup(html, "html.parser")


def parse_document(soup: BeautifulSoup) -> dict:
    """Sayfayi SELECTORS'a gore yapilandirilmis bir sozluge (JSON'a) cevir."""

    # --- Baslik: tek bir eleman -> .select_one ---
    title_el = soup.select_one(SELECTORS["title"])
    title = title_el.get_text(strip=True) if title_el else None

    sections = []

    # --- Bolum 1: Aciklama ---
    desc_el = soup.select_one(SELECTORS["description"])
    if desc_el:
        sections.append({
            "heading": "Description",
            "content": [desc_el.get_text(strip=True)],
        })

    # --- Bolum 2: Urun Bilgisi tablosu (her satir: etiket -> deger) ---
    info_lines = []
    for row in soup.select(SELECTORS["info_rows"]):     # .select -> birden cok eleman dondurur
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

    # --- Her seyi temiz, yapilandirilmis JSON yapisina koy ---
    return {
        "title": title,
        "section_count": len(sections),
        "sections": sections,
    }


def main():
    # Komut satirindan kaynak verilmezse varsayilan URL kullanilir
    source = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_URL
    print(f"[*] Parse ediliyor: {source}")

    soup = load_soup(source)
    document = parse_document(soup)

    # JSON'a cevir. ensure_ascii=False -> Turkce/aksanli karakterler bozulmaz
    output = json.dumps(document, indent=2, ensure_ascii=False)

    # Hem ekrana yaz hem dosyaya kaydet
    print(output)
    with open("document.json", "w", encoding="utf-8") as f:
        f.write(output)
    print("\n[+] document.json yazildi.")


if __name__ == "__main__":
    main()
