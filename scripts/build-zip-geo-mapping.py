"""
Build a compact JSON map of ZIP -> {state, city, county, DMA} for use by
patch-targeting-helper.user.js coverage analysis.

Input:  ../../ZIP_to_DMA_Mapping.xlsx (sibling of repo root, in Downloads)
Output: ../zip-geo-mapping.json

Format (interned to keep size down ~2.5x):

{
  "v": "1",                                 # schema version
  "generated": "YYYY-MM-DD",
  "cities":   ["AGAWAM", "AMHERST", ...],   # unique city strings (preserve original case)
  "counties": ["Hampden County", ...],      # unique county strings
  "dmas":     {"543": "Springfield-Holyoke, MA", ...},
  "zips": {
    "01001": ["MA", 0, 0, "543"],           # [state, cityIdx, countyIdx, dmaCode]
    ...
  }
}

Multi-county zips (All Counties has ";") only use the Primary County for the
forward map. Secondary DMAs are deliberately ignored per product decision.
"""

import json
import sys
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
XLSX = REPO_ROOT.parent / "ZIP_to_DMA_Mapping.xlsx"
OUT = REPO_ROOT / "zip-geo-mapping.json"


def main():
    if not XLSX.exists():
        sys.exit(f"Missing source file: {XLSX}")

    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["ZIP to DMA"]

    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {name: i for i, name in enumerate(header)}

    need = ["ZIP", "State", "USPS Preferred City", "Primary County",
            "Primary DMA Name", "Primary DMA Code"]
    for k in need:
        if k not in idx:
            sys.exit(f"Missing expected column: {k!r}. Got {header!r}")

    cities_index = {}   # city -> idx
    counties_index = {} # county -> idx
    dma_names = {}      # code -> name
    zips = {}

    cities_list = []
    counties_list = []

    def intern(s, table, listref):
        if s in table:
            return table[s]
        i = len(listref)
        table[s] = i
        listref.append(s)
        return i

    seen = 0
    skipped = 0
    for row in rows:
        zip_raw = row[idx["ZIP"]]
        if zip_raw is None:
            continue
        zip_str = str(zip_raw).strip()
        # Normalize to 5-digit string
        if zip_str.isdigit():
            zip_str = zip_str.zfill(5)
        if len(zip_str) != 5 or not zip_str.isdigit():
            skipped += 1
            continue

        state = (row[idx["State"]] or "").strip().upper()
        city = (row[idx["USPS Preferred City"]] or "").strip()
        # Title-case city for nicer display (data has some UPPERCASE)
        if city and city.isupper():
            city = city.title()
        county = (row[idx["Primary County"]] or "").strip()
        dma_name = (row[idx["Primary DMA Name"]] or "").strip()
        dma_code_raw = row[idx["Primary DMA Code"]]
        dma_code = str(dma_code_raw).strip() if dma_code_raw is not None else ""

        if not state or not city or not county or not dma_code:
            skipped += 1
            continue

        city_i = intern(city, cities_index, cities_list)
        county_i = intern(county, counties_index, counties_list)
        if dma_code not in dma_names:
            dma_names[dma_code] = dma_name

        zips[zip_str] = [state, city_i, county_i, dma_code]
        seen += 1

    out = {
        "v": "1",
        "generated": __import__("datetime").date.today().isoformat(),
        "stats": {"zips": seen, "skipped": skipped,
                  "cities": len(cities_list),
                  "counties": len(counties_list),
                  "dmas": len(dma_names)},
        "cities": cities_list,
        "counties": counties_list,
        "dmas": dma_names,
        "zips": zips,
    }

    # Compact JSON: no spaces, sorted keys would inflate output for the zips
    # dict (it's already insertion-sorted by ZIP).
    text = json.dumps(out, separators=(",", ":"), ensure_ascii=False)
    OUT.write_text(text, encoding="utf-8")
    print(f"Wrote {OUT}  ({len(text):,} bytes)  zips={seen} cities={len(cities_list)} counties={len(counties_list)} dmas={len(dma_names)} skipped={skipped}")


if __name__ == "__main__":
    main()
