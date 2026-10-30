"""Debug: print structure of the 3 merged KPI workbooks to the Actions log.

No commits - diagnosis only.
"""
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
EXPORT_DIR = ROOT / "data" / "exports"

FILES = [
    "accounts_opened_merged.xlsx",
    "accounts_closed_merged.xlsx",
    "net_addition_merged.xlsx",
]


def log(msg):
    print(msg, flush=True)


def main():
    for fname in FILES:
        f = EXPORT_DIR / fname
        log("=" * 80)
        log(f"FILE: {fname}")
        if not f.exists():
            log("  MISSING")
            continue
        xls = pd.ExcelFile(f)
        log(f"  Sheets: {xls.sheet_names}")
        for sheet in xls.sheet_names:
            df = pd.read_excel(f, sheet_name=sheet, header=None)
            log(f"  Sheet '{sheet}': shape={df.shape}")
            log("  First 12 rows (repr):")
            for i, row in enumerate(df.head(12).itertuples(index=False, name=None)):
                log(f"    row {i}: {[repr(c) for c in row]}")


if __name__ == "__main__":
    main()
