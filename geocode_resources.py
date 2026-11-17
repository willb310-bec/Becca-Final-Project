#!/usr/bin/env python3
"""
Geocode the BKCB16 resource list with NYC GeoSearch (NYC Planning Labs).

Usage:
    pip install requests openpyxl
    python geocode_resources.py BKCB16_Resources.xlsx

    # preview the cleaned queries without calling the API:
    python geocode_resources.py BKCB16_Resources.xlsx --dry-run

Outputs (next to the input file):
    BKCB16_Resources_geocoded.xlsx   original sheet + Latitude, Longitude,
                                     Matched Address, Confidence, Geocode Status
    BKCB16_Resources_geocoded.csv    same data, ready for QGIS / Google My Maps

Geocode Status values:
    OK                 matched with confidence >= threshold
    LOW_CONFIDENCE     matched, but check the Matched Address column
    BOROUGH_MISMATCH   matched in a different borough than the address says
    NO_MATCH           the geocoder returned nothing
    SKIPPED_PARK_AREA  parks given as cross streets / blocks (not a point address)
    SKIPPED_PO_BOX     P.O. Box, no physical location
    ERROR              request failed after retries
"""
import argparse
import csv
import re
import sys
import time
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

API_URL = "https://geosearch.planninglabs.nyc/v2/search"
# Bias results toward Brownsville so ambiguous street names resolve locally.
FOCUS_LAT, FOCUS_LON = 40.6625, -73.9136
MIN_CONFIDENCE = 0.8          # below this -> LOW_CONFIDENCE
PAUSE_SECONDS = 0.15          # be polite to the free public API
HEADERS = {"User-Agent": "bkcb16-resource-geocoder/1.0"}

BOROUGHS = ["Brooklyn", "Manhattan", "Bronx", "Staten Island", "Queens"]
NUMBER_WORDS = {"one": "1", "two": "2", "three": "3"}


def expected_borough(address: str):
    """Borough named in the address text, if any."""
    for b in BOROUGHS:
        if re.search(rf"\b{b}\b", address, re.I):
            return b
    return None


def clean_address(address: str) -> str:
    """Make a human-written address friendlier to the geocoder."""
    a = address.strip()
    # Drop suite / floor / unit text: ", Suite 1", ", 3rd floor", " #3A", " Suit #348"
    a = re.sub(r",?\s*\b(suites?|suit|ste|unit|apt)\b\.?\s*#?\s*\w+(\s+and\s+\w+)?",
               "", a, flags=re.I)
    # Source typo: "1420 Bushwick Suit #348" is 1420 Bushwick Avenue
    a = re.sub(r"\bBushwick\b(?!\s+Ave)", "Bushwick Avenue", a, flags=re.I)
    a = re.sub(r",?\s*\d+(st|nd|rd|th)\s+floor", "", a, flags=re.I)
    a = re.sub(r",?\s*(first|second|third)\s+floor", "", a, flags=re.I)
    a = re.sub(r"\s+#\s*\w+", "", a)
    # Leading number words: "One Brookdale Plaza" -> "1 Brookdale Plaza"
    m = re.match(r"^(one|two|three)\s+", a, re.I)
    if m:
        a = NUMBER_WORDS[m.group(1).lower()] + " " + a[m.end():]
    # Address ranges: "1746-60 Broadway" / "411-425 Howard Avenue" -> first number
    a = re.sub(r"^(\d+)-\d+\b", r"\1", a)
    # Queens-style hyphen numbers like "20-22 Radde Place" are handled by the rule
    # above (first number). Trailing unit letters: "134A Newport" -> "134 Newport"
    a = re.sub(r"^(\d+)[A-Za-z]\b", r"\1", a)
    # Missing space typo: "1721Pitkin" -> "1721 Pitkin"
    a = re.sub(r"^(\d+)([A-Za-z]{3,})", r"\1 \2", a)
    return re.sub(r"\s+", " ", a).strip(" ,")


def geocode(session: requests.Session, query: str, retries: int = 3):
    params = {
        "text": query,
        "size": 1,
        "focus.point.lat": FOCUS_LAT,
        "focus.point.lon": FOCUS_LON,
    }
    last_err = None
    for attempt in range(retries):
        try:
            r = session.get(API_URL, params=params, headers=HEADERS, timeout=20)
            if r.status_code == 429:
                time.sleep(2 * (attempt + 1))
                continue
            r.raise_for_status()
            feats = r.json().get("features", [])
            return feats[0] if feats else None
        except (requests.RequestException, ValueError) as e:
            last_err = e
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"request failed: {last_err}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="BKCB16_Resources.xlsx")
    ap.add_argument("--sheet", default="Resources")
    ap.add_argument("--dry-run", action="store_true",
                    help="print cleaned queries only; no API calls")
    args = ap.parse_args()

    src = Path(args.input)
    wb = load_workbook(src)
    ws = wb[args.sheet]

    header = [c.value for c in ws[1]]
    col = {name: i for i, name in enumerate(header)}
    for needed in ("Organization", "Address", "Address Type"):
        if needed not in col:
            sys.exit(f"Missing column: {needed}")

    new_cols = ["Latitude", "Longitude", "Matched Address", "Confidence", "Geocode Status"]
    start = len(header) + 1
    for i, name in enumerate(new_cols):
        cell = ws.cell(row=1, column=start + i, value=name)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="1F3864")

    session = requests.Session()
    cache = {}
    stats = {}
    csv_rows = [header + new_cols]

    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row=row, column=col["Address"] + 1).value or ""
        atype = ws.cell(row=row, column=col["Address Type"] + 1).value or ""
        lat = lon = label = conf = None

        if atype == "Cross streets":
            status = "SKIPPED_PARK_AREA"
        elif atype == "P.O. Box" or raw.upper().startswith("P.O. BOX"):
            status = "SKIPPED_PO_BOX"
        else:
            query = clean_address(raw)
            if args.dry_run:
                print(f"{raw!r:70} -> {query!r}")
                status = "DRY_RUN"
            else:
                try:
                    if query not in cache:
                        cache[query] = geocode(session, query)
                        time.sleep(PAUSE_SECONDS)
                    feat = cache[query]
                except RuntimeError as e:
                    feat, status = None, "ERROR"
                    print(f"row {row}: {e}", file=sys.stderr)
                else:
                    if not feat:
                        status = "NO_MATCH"
                    else:
                        lon, lat = feat["geometry"]["coordinates"]
                        props = feat.get("properties", {})
                        label = props.get("label")
                        conf = props.get("confidence")
                        want = expected_borough(raw)
                        got = props.get("borough")
                        if want and got and want.lower() != got.lower():
                            status = "BOROUGH_MISMATCH"
                        elif conf is not None and conf < MIN_CONFIDENCE:
                            status = "LOW_CONFIDENCE"
                        else:
                            status = "OK"

        stats[status] = stats.get(status, 0) + 1
        values = [lat, lon, label, conf, status]
        for i, v in enumerate(values):
            c = ws.cell(row=row, column=start + i, value=v)
            c.font = Font(name="Arial", size=10)
            if i in (0, 1) and v is not None:
                c.number_format = "0.000000"
        csv_rows.append([ws.cell(row=row, column=j + 1).value for j in range(len(header))] + values)

    for i, w in enumerate([12, 12, 50, 11, 20]):
        ws.column_dimensions[ws.cell(row=1, column=start + i).column_letter].width = w
    ws.auto_filter.ref = ws.dimensions

    if args.dry_run:
        print("\nDry run only; no files written.")
        return

    out_xlsx = src.with_name(src.stem + "_geocoded.xlsx")
    out_csv = src.with_name(src.stem + "_geocoded.csv")
    wb.save(out_xlsx)
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(csv_rows)

    print("\nResults:")
    for k, v in sorted(stats.items()):
        print(f"  {k:20} {v}")
    print(f"\nSaved {out_xlsx} and {out_csv}")
    print("Review rows whose Geocode Status is not OK before mapping.")


if __name__ == "__main__":
    main()
