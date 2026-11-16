#!/usr/bin/env python3
"""
Haalt bekendmakingen van de Gemeente Den Haag op via de officiële SRU-API van
KOOP (dezelfde data als zoek.officielebekendmakingen.nl) en voegt nieuwe
bekendmakingen toe aan een Excel-bestand. Reeds aanwezige bekendmakingen
(op basis van het identifier-nummer, bv. gmb-2026-123456) worden overgeslagen.

Uitvoer:
  bekendmakingen_den_haag.xlsx     alle gegevens
  bekendmakingen_den_haag.csv      zelfde, als tabel te bekijken op GitHub
  bekendmakingen_den_haag.geojson  punten op een kaart (GitHub toont dit als kaart)

Vereist:   pip install requests openpyxl shapely pyproj
"""
import csv
import json
import re
import sys
import datetime as dt
import xml.etree.ElementTree as ET
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pyproj import Transformer
from shapely import wkt as swkt
from shapely.geometry import GeometryCollection, mapping
from shapely.ops import transform as shp_transform

# ----------------------------------------------------------------- instellingen
SRU_URL = "https://repository.overheid.nl/sru"
# Mogelijke schrijfwijzen van de organisatienaam in de API; de eerste met resultaten wint
CREATORS = ["'s-Gravenhage", "Den Haag", "Gemeente Den Haag", "Gemeente 's-Gravenhage"]
DAYS_BACK = 30                 # kijk een paar dagen terug; dubbelen worden gefilterd
MAP_FILTER = "vergunning"     # alleen deze typen op de kaart (deel van het type); "" = alles
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
    ("Latitude", 12),
    ("Longitude", 12),
    ("Geometrie (WGS84 WKT)", 40),
]
NAMES = [c for c, _ in COLUMNS]
WRAP = {"Titel", "Adres", "Beschrijving"}

RD_TO_WGS = Transformer.from_crs("EPSG:28992", "EPSG:4326", always_xy=True)


# ------------------------------------------------------------------- ophalen
def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _query(query: str):
    """Voert de zoekopdracht uit (met paginering). Geeft (records, totaal, eerste_respons)."""
    records, start, total, first = [], 1, 0, b""
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
        if start == 1:
            first = r.content
        root = ET.fromstring(r.content)
        page = [e for e in root.iter() if local(e.tag) == "record"]
        records.extend(page)
        total = next((int(e.text) for e in root.iter() if local(e.tag) == "numberOfRecords"), 0)
        start += PAGE_SIZE
        if not page or start > total:
            break
    return records, total, first


def fetch_records(since: dt.date):
    """Probeert de organisatienamen één voor één; geeft de records van de eerste met resultaat."""
    last_response = b""
    for naam in CREATORS:
        query = (
            f'c.product-area==officielepublicaties AND dt.creator=="{naam}" '
            f'AND dt.available>={since.isoformat()}'
        )
        records, total, first = _query(query)
        print(f"Zoekopdracht: {query}\n  -> {total} resultaten")
        last_response = first
        if records:
            return records
    print("GEEN resultaten voor alle namen. Begin van de laatste API-respons:")
    print(last_response[:2000].decode("utf-8", errors="replace"))
    try:
        r = requests.get(SRU_URL, params={"operation": "scan", "version": "2.0",
                                          "scanClause": "dt.creator", "maximumTerms": 2000},
                         timeout=TIMEOUT)
        namen = sorted({e.text for e in ET.fromstring(r.content).iter()
                        if local(e.tag) == "value" and e.text and "haag" in e.text.lower()})
        print("Organisaties met 'haag' in de naam:", namen or "(scan gaf niets terug)")
    except Exception as e:
        print("Scan mislukt:", e)
    return []


# ------------------------------------------------------------------- geometrie
WKT_START = re.compile(r"^(SRID=\d+;)?\s*(POINT|LINESTRING|POLYGON|MULTIPOINT|MULTILINESTRING|"
                       r"MULTIPOLYGON|GEOMETRYCOLLECTION)", re.I)
_geo_samples = []   # voor diagnose in het logboek


def to_wgs84(geom):
    """Zet een geometrie om naar lengte/breedte (x=lon, y=lat); herkent RD en ETRS89 (ook omgekeerd)."""
    minx, miny, maxx, maxy = geom.bounds
    if maxx > 1000 or maxy > 1000:                       # Rijksdriehoek (meters)
        geom = shp_transform(RD_TO_WGS.transform, geom)
    elif 50 <= minx <= 54 and 3 <= miny <= 8:            # (lat, lon) volgorde -> omwisselen
        geom = shp_transform(lambda x, y, z=None: (y, x), geom)
    return geom


def extract_geometry(rec):
    """Zoekt WKT-geometrieën in het record. Geeft één WGS84-geometrie of None."""
    geoms, seen = [], set()
    for el in rec.iter():
        t = (el.text or "").strip()
        if not t or not WKT_START.match(t) or t in seen:
            continue
        seen.add(t)
        if len(_geo_samples) < 2:
            _geo_samples.append(f"{local(el.tag)}: {t[:200]}")
        try:
            g = to_wgs84(swkt.loads(re.sub(r"^SRID=\d+;", "", t, flags=re.I)))
            minx, miny, maxx, maxy = g.bounds
            if 3 <= minx and maxx <= 8 and 50 <= miny and maxy <= 54:   # ligt in Nederland
                geoms.append(g)
        except Exception:
            continue
    if not geoms:
        return None
    return geoms[0] if len(geoms) == 1 else GeometryCollection(geoms)


def wkt_for_cell(geom) -> str:
    for tol in (0, 0.0001, 0.001):
        g = geom.simplify(tol) if tol else geom
        s = swkt.dumps(g, rounding_precision=6)
        if len(s) < 32000:
            return s
    return ""


# ------------------------------------------------------------------- parsen
def parse_record(rec) -> dict:
    vals: dict[str, list[str]] = {}
    for el in rec.iter():
        txt = (el.text or "").strip()
        if txt:
            vals.setdefault(local(el.tag), [])
            if txt not in vals[local(el.tag)]:
                vals[local(el.tag)].append(txt)

    def get(*names, sep="; "):
        for n in names:
            if n in vals:
                return sep.join(vals[n])
        return ""

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

    d = {
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
        "Link": get("preferredUrl", "prefferedUrl", "url"),
        "Identifier": get("identifier"),
        "Latitude": "", "Longitude": "", "Geometrie (WGS84 WKT)": "",
    }
    geom = extract_geometry(rec)
    if geom is not None:
        c = geom.centroid
        d["Latitude"], d["Longitude"] = round(c.y, 6), round(c.x, 6)
        d["Geometrie (WGS84 WKT)"] = wkt_for_cell(geom)
    return d


# -------------------------------------------------------------------- excel
def style_row(ws, r):
    for i, name in enumerate(NAMES, 1):
        c = ws.cell(row=r, column=i)
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(vertical="top", wrap_text=name in WRAP)
    lc = ws.cell(row=r, column=NAMES.index("Link") + 1)
    if lc.value:
        lc.hyperlink = lc.value
        lc.font = Font(name="Arial", size=10, color="0563C1", underline="single")


def style_header(ws, i):
    cell = ws.cell(row=1, column=i)
    cell.font = Font(name="Arial", bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.column_dimensions[get_column_letter(i)].width = COLUMNS[i - 1][1]


def open_workbook():
    if OUTPUT.exists():
        wb = load_workbook(OUTPUT)
        ws = wb.active
        # Oudere bestanden: ontbrekende kolommen achteraan toevoegen
        have = [c.value for c in ws[1]]
        for name in NAMES:
            if name not in have:
                ws.cell(row=1, column=len(have) + 1, value=name)
                have.append(name)
        return wb, ws
    wb = Workbook()
    ws = wb.active
    ws.title = "Bekendmakingen"
    ws.append(NAMES)
    for i in range(1, len(NAMES) + 1):
        style_header(ws, i)
    ws.freeze_panes = "A2"
    return wb, ws


def write_outputs(ws):
    header = [c.value for c in ws[1]]
    rows = [dict(zip(header, r)) for r in ws.iter_rows(min_row=2, values_only=True)]

    # CSV (alle kolommen, in de volgorde van het werkblad)
    with open(OUTPUT.with_suffix(".csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        for r in rows:
            w.writerow(["" if r.get(h) is None else r.get(h) for h in header])

    # GeoJSON: één punt per bekendmaking met coördinaten
    feats = []
    for r in rows:
        lat, lon = r.get("Latitude"), r.get("Longitude")
        if lat in (None, "") or lon in (None, ""):
            continue
        if MAP_FILTER and MAP_FILTER.lower() not in str(r.get("Type") or "").lower():
            continue
        props = {
            "titel": r.get("Titel"), "type": r.get("Type"), "adres": r.get("Adres"),
            "gepubliceerd": r.get("Datum gepubliceerd"), "link": r.get("Link"),
            "identifier": r.get("Identifier"),
            "marker-color": "#d62728", "marker-size": "medium",
        }
        feats.append({"type": "Feature", "properties": props,
                      "geometry": {"type": "Point", "coordinates": [float(lon), float(lat)]}})
    with open(OUTPUT.with_suffix(".geojson"), "w", encoding="utf-8") as f:
        json.dump({"type": "FeatureCollection", "features": feats}, f, ensure_ascii=False, indent=1)
    return len(feats)


def main() -> int:
    since = dt.date.today() - dt.timedelta(days=DAYS_BACK)
    try:
        records = fetch_records(since)
    except Exception as e:
        print(f"FOUT bij ophalen: {e}", file=sys.stderr)
        return 1

    wb, ws = open_workbook()
    header = [c.value for c in ws[1]]
    id_idx = header.index("Identifier")
    bekend = {str(r[id_idx]) for r in ws.iter_rows(min_row=2, values_only=True) if r[id_idx]}

    nieuw = met_coord = 0
    nu = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for rec in records:
        d = parse_record(rec)
        if not d["Identifier"] or d["Identifier"] in bekend:
            continue
        d["Opgehaald op"] = nu
        ws.append([d.get(h, "") for h in header])
        style_row(ws, ws.max_row)
        bekend.add(d["Identifier"])
        nieuw += 1
        met_coord += d["Latitude"] != ""

    ws.auto_filter.ref = ws.dimensions
    wb.save(OUTPUT)
    punten = write_outputs(ws)
    print(f"{dt.datetime.now():%Y-%m-%d %H:%M} - {len(records)} gevonden, {nieuw} nieuw toegevoegd "
          f"({met_coord} met coördinaten); {punten} punten op de kaart")
    if records and nieuw and not met_coord:
        print("Geen coördinaten herkend. Voorbeelden van geometrie-tekst in de data:",
              _geo_samples or "(geen WKT-velden gevonden)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
