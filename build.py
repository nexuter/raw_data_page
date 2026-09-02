"""Regenerate index.html from db_lits.xlsx.

Usage: python build.py
Reads the spreadsheet next to this script and injects the rows into
template.html at the __DB_DATA__ marker, producing a self-contained index.html.
"""
import json
import pathlib
import re

import openpyxl

HERE = pathlib.Path(__file__).parent
SOURCE = HERE / "db_lits.xlsx"
TEMPLATE = HERE / "template.html"
OUTPUT = HERE / "index.html"

# Titles reached by asking the contact rather than by following the link.
BY_REQUEST = {"Glassdoor", "WRDS"}


def sort_key(name):
    """Library filing order: a leading article does not decide the shelf."""
    return re.sub(r"^(the|a|an)\s+", "", name.strip(), flags=re.I).lower()


def main():
    ws = openpyxl.load_workbook(SOURCE, data_only=True)["Sheet1"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() for h in rows[0]]

    entries = []
    for raw in rows[1:]:
        if not any(raw):
            continue
        rec = {header[i]: (str(raw[i]).strip() if raw[i] is not None else "") for i in range(len(header))}
        contact_url = rec["url-contact"]
        entries.append({
            "name": rec["name"],
            "url": rec["url-name"],
            "summary": rec["summary"],
            "contact": rec["contact"],
            "contactUrl": ("mailto:" + contact_url) if "@" in contact_url and "://" not in contact_url else contact_url,
            "byRequest": rec["name"] in BY_REQUEST,
            "letter": sort_key(rec["name"])[0].upper(),
        })

    entries.sort(key=lambda e: sort_key(e["name"]))
    payload = json.dumps(entries, ensure_ascii=False, indent=1)
    OUTPUT.write_text(TEMPLATE.read_text(encoding="utf-8").replace("__DB_DATA__", payload), encoding="utf-8")
    print(f"{len(entries)} databases -> {OUTPUT.name}")


if __name__ == "__main__":
    main()
