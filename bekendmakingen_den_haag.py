#!/usr/bin/env python3
"""
Haalt bekendmakingen van de Gemeente Den Haag op via de officiële SRU-API van
KOOP (dezelfde data als zoek.officielebekendmakingen.nl) en voegt nieuwe
bekendmakingen toe aan een Excel-bestand. Reeds aanwezige bekendmakingen
(op basis van het identifier-nummer, bv. gmb-2026-123456) worden overgeslagen.

Gebruik:   python bekendmakingen_den_haag.py
Vereist:   pip install requests openpyxl
"""
import sys
import datetime as dt
import xml.etree.ElementTree as ET
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------- instellingen
SRU_URL = "https://repository.overheid.nl/sru"
CREATOR = "Den Haag"          # waarde van dt.creator; controleer evt. met scanClause=dt.creator
DAYS_BACK = 3                 # kijk een paar dagen terug; dubbelen worden gefilterd
OUTPUT = Path(__file__).with_name("bekendmakingen_den_haag.xlsx")
PAGE_SIZE = 500
TIMEOUT = 60

COLUMNS = [
    ("Datum gepubliceerd", 16),
    ("Type", 28),
    ("Titel", 60),
    ("Adres", 40),
    ("Postcode", 12),
    ("Woonplaats", 16),
    ("Beschrijving", 70),
    ("Startdatum", 13),
    ("Einddatum", 13),
    ("Einde reactietermijn", 16),
    ("Publicatieblad", 16),
    ("Link", 50),
    ("Identifier", 22),
    ("Opgehaald op", 18),
]
ID_COL = COLUMNS.index(("Identifier", 22)) + 1


# ------------------------------------------------------------------- ophalen
def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def fetch_records(since: dt.date):
    """Geeft een lijst met <record>-elementen sinds de opgegeven datum."""
    query = (
        f'c.product-area==officielepublicaties AND dt.creator=="{CREATOR}" '
        f'AND dt.available>={since.isoformat()}'
    )
    records, start = [], 1
    while True:
        r = requests.get(
            SRU_URL,
            params={
                "query": query,
                "startRecord": start,
                "maximumRecords": PAGE_SIZE,
                "version": "2.0",
                "operation": "searchRetrieve",
            },
            timeout=TIMEOUT,
            headers={"User-Agent": "bekendmakingen-export/1.0"},
        )
        r.raise_for_status()
        root = ET.fromstring(r.content)
        page = [e for e in root.iter() if local(e.tag) == "record"]
        records.extend(page)
        total = next((int(e.text) for e in root.iter() if local(e.tag) == "numberOfRecords"), 0)
        start += PAGE_SIZE
        if not page or start > total:
            break
    return records


# ------------------------------------------------------------------- parsen
def parse_record(rec) -> dict:
    """Verzamelt alle velden per lokale naam; meerdere waarden worden uniek bewaard."""
    vals: dict[str, list[str]] = {}
    for el in rec.iter():
        txt = (el.text or "").strip()
        if txt:
            vals.setdefault(local(el.tag), [])
            if txt not in vals[local(el.tag)]:
                vals[local(el.tag)].append(txt)
        # URL staat soms in een attribuut (bv. itemUrl / hasVersion)
    def get(*names, sep="; "):
        for n in names:
            if n in vals:
                return sep.join(vals[n])
        return ""

    # Adres: straat + huisnummer(+letter/toevoeging), meerdere locaties worden gejoined
    straten = vals.get("straatnaam", [])
    nrs = vals.get("huisnummer", [])
    letters = vals.get("huisletter", [])
    toev = vals.get("huisnummertoevoeging", [])
    if straten:
        adres = "; ".join(straten)
        if nrs:
            adres += " " + ", ".join(
                n + (letters[i] if i < len(letters) else "") + (("-" + toev[i]) if i < len(toev) else "")
                for i, n in enumerate(nrs)
            )
    else:
        adres = get("geometrielabel", "postcodeHuisnummer", "spatial")

    link = get("preferredUrl", "prefferedUrl", "url")

    return {
        "Datum gepubliceerd": get("available", "issued", "modified"),
        "Type": get("type"),
        "Titel": get("title"),
        "Adres": adres,
        "Postcode": get("postcode"),
        "Woonplaats": get("woonplaats", "gemeentenaam"),
        "Beschrijving": get("abstract", "alternative"),
        "Startdatum": get("startdatum"),
        "Einddatum": get("einddatum"),
        "Einde reactietermijn": get("datumEindeReactietermijn"),
        "Publicatieblad": get("publicatienaam"),
        "Link": link,
        "Identifier": get("identifier"),
    }


# -------------------------------------------------------------------- excel
def open_workbook():
    if OUTPUT.exists():
        wb = load_workbook(OUTPUT)
        return wb, wb.active
    wb = Workbook()
    ws = wb.active
    ws.title = "Bekendmakingen"
    ws.append([c for c, _ in COLUMNS])
    for i, (_, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=i)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    return wb, ws


def main() -> int:
    since = dt.date.today() - dt.timedelta(days=DAYS_BACK)
    try:
        records = fetch_records(since)
    except Exception as e:  # netwerk/parse-fout
        print(f"FOUT bij ophalen: {e}", file=sys.stderr)
        return 1

    wb, ws = open_workbook()
    bekend = {str(r[ID_COL - 1]) for r in ws.iter_rows(min_row=2, values_only=True) if r[ID_COL - 1]}

    nieuw = 0
    nu = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for rec in records:
        d = parse_record(rec)
        if not d["Identifier"] or d["Identifier"] in bekend:
            continue
        row = [d.get(c, "") for c, _ in COLUMNS[:-1]] + [nu]
        ws.append(row)
        r = ws.max_row
        for i in range(1, len(COLUMNS) + 1):
            c = ws.cell(row=r, column=i)
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=i in (3, 4, 7))
        link_cell = ws.cell(row=r, column=COLUMNS.index(("Link", 50)) + 1)
        if link_cell.value:
            link_cell.hyperlink = link_cell.value
            link_cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")
        bekend.add(d["Identifier"])
        nieuw += 1

    ws.auto_filter.ref = ws.dimensions
    wb.save(OUTPUT)
    print(f"{dt.datetime.now():%Y-%m-%d %H:%M} - {len(records)} gevonden, {nieuw} nieuw toegevoegd -> {OUTPUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
