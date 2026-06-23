"""
Patent XML Parser - IPPH/WIPO-style XML -> Structured JSON (Prototype)
======================================================================
Parses Chinese patent data (IPPH format, schemaVersion V2.1.0) XML files into
clean, structured output. Package layout:

    <date-folder>/CREATE_xxx.ZIP  ->  inside the zip  ->  CN<...>A.XML

This prototype parses a single patent XML (either a file path directly, or the
first XML inside a zip). Extracted fields match the requested scope:
  - publication / application identifiers
  - claims + claim numbers + independent/dependent indicator
  - description sections
  - bibliographic metadata (IPC classifications, dates)
  - legal/owner metadata (applicant / inventor)

Usage:
    pip install lxml
    python patent_xml_parser.py <patent.xml>      # a single XML file
    python patent_xml_parser.py <package.zip>      # the first XML inside a zip
"""

import sys
import json
import zipfile

from lxml import etree


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _text(el):
    """Return an element's whitespace-normalized text (None-safe)."""
    if el is None:
        return None
    txt = "".join(el.itertext())
    return " ".join(txt.split()) or None


def load_root(source: str):
    """Parse the first XML inside a .zip, or the file itself otherwise.
    recover=True lets us parse despite the external DTD/entity reference."""
    parser = etree.XMLParser(recover=True, resolve_entities=False, no_network=True)
    if source.lower().endswith(".zip"):
        with zipfile.ZipFile(source) as zf:
            name = next(n for n in zf.namelist() if n.lower().endswith(".xml"))
            data = zf.read(name)
        return etree.fromstring(data, parser=parser)
    with open(source, "rb") as f:
        return etree.fromstring(f.read(), parser=parser)


# ---------------------------------------------------------------------------
# Field extractors  (each maps to explicit tags/attributes in the source XML)
# ---------------------------------------------------------------------------

def extract_identifiers(doc):
    """Publication and application IDs (multiple dataFormats: original/ipph/docdb/epodoc)."""
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
    """Claim blocks: summary counts + each claim (number, independent/dependent, parent ref, text)."""
    blocks = []
    for claims in doc.findall(".//Claims"):
        items = []
        for cl in claims.findall("Claim"):
            items.append({
                "id": cl.get("id"),
                "number": cl.get("number"),
                "subordination": cl.get("subordination"),   # independent | sub
                "is_independent": cl.get("subordination") == "independent",
                "parent_ref": cl.get("idRef"),               # parent claim a dependent claim refers to
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
    """Invention title + description paragraphs + drawings description."""
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
    """IPC classification codes (bibliographic metadata)."""
    return [_text(c.find("Text")) for c in doc.findall(".//ClassificationIPCR")
            if _text(c.find("Text"))]


def extract_parties(doc):
    """Legal/owner metadata: applicants and inventors (duplicate names deduplicated)."""
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
    """Convert one patent XML root into a fully structured dict."""
    doc = root.find(".//PatentDocument")
    if doc is None:
        doc = root
    ids = extract_identifiers(doc)
    # Surface the primary publication number (ipph format) for easy access
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
        print("Usage: python patent_xml_parser.py <patent.xml | package.zip>")
        sys.exit(1)
    source = sys.argv[1]
    print(f"[*] Parsing: {source}")

    root = load_root(source)
    record = parse_patent(root)

    output = json.dumps(record, indent=2, ensure_ascii=False)
    with open("patent.json", "w", encoding="utf-8") as f:
        f.write(output)

    # Summary (quick indicator for validator / manual review)
    c = record["claims"][0] if record["claims"] else {}
    print(f"[+] Publication: {record['primary_publication_id']}")
    print(f"[+] Title: {record['description']['invention_title']}")
    print(f"[+] Claims: {c.get('claims_quantity')} (independent={c.get('independent_quantity')}, sub={c.get('sub_quantity')})")
    print(f"[+] IPC: {', '.join(record['ipc_classifications'][:5])}")
    print(f"[+] Applicants: {record['parties']['applicants']}")
    print(f"[+] Description paragraphs: {record['description']['paragraph_count']}")
    print("[+] Wrote patent.json (full structured output).")


if __name__ == "__main__":
    main()
