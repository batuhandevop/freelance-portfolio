"""
Patent XML Parser - IPPH/WIPO-style XML -> Structured JSON (Prototype)
======================================================================
Cin patent verisi (IPPH formati, schemaVersion V2.1.0) XML'lerini temiz,
yapilandirilmis cikti'ya cevirir. Veri paketi yapisi:

    <gun-klasoru>/CREATE_xxx.ZIP  ->  zip icinde  ->  CN<...>A.XML

Bu prototip TEK bir patent XML'ini parse eder (dogrudan dosyadan ya da bir
zip girisinden). Cikan alanlar musterinin istedikleridir:
  - publication / application identifiers
  - claims + claim numbers + independent/dependent gostergesi
  - description bolumleri
  - bibliographic metadata (IPC siniflandirma, tarihler)
  - legal/owner metadata (applicant / inventor)

Calistirma:
    pip install lxml
    python patent_xml_parser.py <patent.xml>      # tek XML dosyasi
    python patent_xml_parser.py <paket.zip>        # zip icindeki ilk XML
"""

import sys
import json
import zipfile

from lxml import etree


# ---------------------------------------------------------------------------
# Yardimcilar
# ---------------------------------------------------------------------------

def _text(el):
    """Bir elemanin metnini bosluklar temizlenmis dondur (None-guvenli)."""
    if el is None:
        return None
    txt = "".join(el.itertext())
    return " ".join(txt.split()) or None


def load_root(source: str):
    """Kaynak .zip ise icindeki ilk XML'i, degilse dosyayi parse eder.
    recover=True -> harici DTD/entity referanslarina takilmadan parse eder."""
    parser = etree.XMLParser(recover=True, resolve_entities=False, no_network=True)
    if source.lower().endswith(".zip"):
        with zipfile.ZipFile(source) as zf:
            name = next(n for n in zf.namelist() if n.lower().endswith(".xml"))
            data = zf.read(name)
        return etree.fromstring(data, parser=parser)
    with open(source, "rb") as f:
        return etree.fromstring(f.read(), parser=parser)


# ---------------------------------------------------------------------------
# Alan cikariciler  (her biri kaynak XML'deki acik tag/attribute'lara isaret eder)
# ---------------------------------------------------------------------------

def extract_identifiers(doc):
    """publication ve application ID'leri (birden cok dataFormat var: original/ipph/docdb/epodoc)."""
    def collect(tag):
        out = []
        for info in doc.findall(f".//BibliographicData/{tag}"):
            did = info.find("DocumentID")
            if did is None:
                continue
            out.append({
                "data_format": info.get("dataFormat"),
                "dnum": did.get("dnum"),
                "country": _text(did.find("Country")),
                "number": _text(did.find("Number")),
                "kind": _text(did.find("Kind")),
                "date": _text(did.find("Date")),
            })
        return out
    return {
        "publication": collect("PublicationInfo"),
        "application": collect("ApplicationInfo"),
    }


def extract_claims(doc):
    """Claims bloklari: ozet sayilar + her claim (number, independent/dependent, parent ref, metin)."""
    blocks = []
    for claims in doc.findall(".//Claims"):
        items = []
        for cl in claims.findall("Claim"):
            items.append({
                "id": cl.get("id"),
                "number": cl.get("number"),
                "subordination": cl.get("subordination"),   # independent | sub
                "is_independent": cl.get("subordination") == "independent",
                "parent_ref": cl.get("idRef"),               # bagimli claim'in baglandigi claim
                "text": " ".join(_text(t) or "" for t in cl.findall("ClaimText")).strip(),
            })
        blocks.append({
            "lang": claims.get("lang"),
            "claims_quantity": claims.get("claimsQuantity"),
            "independent_quantity": claims.get("independentQuantity"),
            "sub_quantity": claims.get("subQuantity"),
            "claims": items,
        })
    return blocks


def extract_description(doc):
    """Bulus basligi + aciklama paragraflari + cizim aciklamasi."""
    desc = doc.find(".//Description")
    paragraphs = [_text(p) for p in desc.findall(".//Paragraph")] if desc is not None else []
    return {
        "invention_title": _text(doc.find(".//InventionTitle")),
        "abstract": _text(doc.find(".//Abstract")),
        "drawings_description": _text(doc.find(".//DrawingsDescription")),
        "paragraph_count": len([p for p in paragraphs if p]),
        "paragraphs": [p for p in paragraphs if p],
    }


def extract_classifications(doc):
    """IPC siniflandirma kodlari (bibliographic metadata)."""
    return [_text(c.find("Text")) for c in doc.findall(".//ClassificationIPCR")
            if _text(c.find("Text"))]


def extract_parties(doc):
    """Legal/owner metadata: applicant'lar ve inventor'lar (tekrarli isimler tekillestirilir)."""
    def names(path):
        seen, out = set(), []
        for n in doc.findall(path):
            name = _text(n)
            if name and name not in seen:
                seen.add(name)
                out.append(name)
        return out
    return {
        "applicants": names(".//Applicants/Applicant/Name"),
        "inventors": names(".//Inventors/Inventor/Name"),
    }


def parse_patent(root) -> dict:
    """Bir patent XML kokunu tam yapilandirilmis sozluge cevir."""
    doc = root.find(".//PatentDocument")
    if doc is None:
        doc = root
    ids = extract_identifiers(doc)
    # Birincil yayin numarasini (ipph format) kolay erisim icin yukari tasi
    pub_ipph = next((p["dnum"] for p in ids["publication"] if p["data_format"] == "ipph"), None)
    return {
        "primary_publication_id": pub_ipph,
        "identifiers": ids,
        "description": extract_description(doc),
        "claims": extract_claims(doc),
        "ipc_classifications": extract_classifications(doc),
        "parties": extract_parties(doc),
    }


def main():
    if len(sys.argv) < 2:
        print("Kullanim: python patent_xml_parser.py <patent.xml | paket.zip>")
        sys.exit(1)
    source = sys.argv[1]
    print(f"[*] Parse ediliyor: {source}")

    root = load_root(source)
    record = parse_patent(root)

    output = json.dumps(record, indent=2, ensure_ascii=False)
    with open("patent.json", "w", encoding="utf-8") as f:
        f.write(output)

    # Ozet (validator/manuel kontrol icin hizli gosterge)
    c = record["claims"][0] if record["claims"] else {}
    print(f"[+] Publication: {record['primary_publication_id']}")
    print(f"[+] Title: {record['description']['invention_title']}")
    print(f"[+] Claims: {c.get('claims_quantity')} (independent={c.get('independent_quantity')}, sub={c.get('sub_quantity')})")
    print(f"[+] IPC: {', '.join(record['ipc_classifications'][:5])}")
    print(f"[+] Applicants: {record['parties']['applicants']}")
    print(f"[+] Description paragraphs: {record['description']['paragraph_count']}")
    print("[+] patent.json yazildi (tam yapilandirilmis cikti).")


if __name__ == "__main__":
    main()
