#!/usr/bin/env python3
"""Regenerate bo_master.js from the Haryana Circle BO master sheet.

    python3 automation/gen_bo_master.py path/to/BO_Master_data_of_Haryana_Circle.xlsx

Reads the `Consolidate` sheet and writes `const BO_MASTER=[{name, code, so,
subDivision, division}, ...]` at the repo root. The output is a valid JSON array,
so `cept_mis_cbs.py` parses the same file. Identity is `code` (SOL_ID_BO_ID): the
sheet has 2,196 offices but only 2,093 distinct names, so names alone cannot tell
two offices apart. Columns the dashboard does not use (targets, Rural/Urban,
single/double hand) are deliberately left out.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SHEET = "Consolidate"
COLS = {
    "name": "Office_Name",
    "code": "SOL_ID_BO_ID",
    "so": "AO_Name",
    "subDivision": "Sub_Division",
    "division": "Division",
}


def cell(row, index, header):
    i = index.get(header)
    v = row[i] if i is not None and i < len(row) else None
    return str(v).strip() if v is not None else ""


def main(src):
    wb = openpyxl.load_workbook(src, data_only=True)
    rows = list(wb[SHEET].iter_rows(values_only=True))
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    index = {h: i for i, h in enumerate(header) if h}
    for key, header_name in COLS.items():
        if header_name not in index:
            raise SystemExit(f"{SHEET} sheet me column '{header_name}' nahi mila (headers: {header[:6]}...)")

    offices = []
    for r in rows[1:]:
        if not any(v is not None and str(v).strip() for v in r):
            continue
        code = cell(r, index, COLS["code"])
        if code == "0":  # the sheet carries one office with code 0 - store it as blank
            code = ""
        offices.append({
            "name": cell(r, index, COLS["name"]),
            "code": code.upper(),
            "so": cell(r, index, COLS["so"]),
            "subDivision": re.sub(r"\s+", " ", cell(r, index, COLS["subDivision"])).strip(),
            "division": re.sub(r"\s+", " ", cell(r, index, COLS["division"])).strip(),
        })

    seen = {}
    for o in offices:
        if o["code"]:
            if o["code"] in seen:
                print(f"WARNING duplicate code {o['code']}: {seen[o['code']]} / {o['name']}")
            seen[o["code"]] = o["name"]
    offices.sort(key=lambda o: (o["name"].lower(), o["code"]))

    js = "const BO_MASTER=" + json.dumps(offices, separators=(",", ":"), ensure_ascii=False) + ";\n"
    (ROOT / "bo_master.js").write_text(js, encoding="utf-8")
    names = {re.sub(r"\s*B\.?O\.?$", " BO", o["name"], flags=re.I).strip().lower() for o in offices if o["name"]}
    print(f"bo_master.js: {len(offices)} offices, {len(seen)} codes, {len(names)} distinct names, "
          f"{len({o['division'] for o in offices})} divisions, "
          f"{len({o['subDivision'] for o in offices if o['subDivision']})} sub-divisions, "
          f"{len(js.encode('utf-8'))} bytes")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    main(sys.argv[1])
