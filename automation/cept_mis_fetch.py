"""
Credential-free CEPT MIS BO report fetcher.

Downloads the BO-wise report from REPORT_URL (no login required),
consolidates it into data/dashboard_data.xlsx for the dashboard.

Usage (GitHub Actions sets REPORT_URL automatically):
    REPORT_URL=https://... python cept_mis_fetch.py
"""
import io, os, sys
from datetime import date
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

REPORT_URL = os.getenv("REPORT_URL", "")
if not REPORT_URL:
    sys.exit("REPORT_URL not set. Add it in repo Settings > Secrets and variables > Actions > Variables.")


def main():
    r = requests.get(REPORT_URL, timeout=60)
    r.raise_for_status()
    content = r.content
    name = REPORT_URL.split("?")[0].lower()
    if name.endswith(".csv") or b"," in content[:200] and b"<html" not in content[:200].lower():
        df = pd.read_csv(io.BytesIO(content))
    else:
        df = pd.read_excel(io.BytesIO(content))
    df.columns = [str(c).strip() for c in df.columns]
    date_col = next((c for c in df.columns if "date" in c.lower()), None)
    bo_col = next((c for c in df.columns if c.lower() in ("bo", "bo name", "office_name", "office name") or "branch" in c.lower()), None)
    if not date_col or not bo_col:
        sys.exit(f"Date/BO column not found. Columns: {list(df.columns)}")
    df = df.rename(columns={date_col: "Date", bo_col: "BO"})
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["BO"] = df["BO"].astype(str).str.strip()
    keep = ["Date", "BO"] + [c for c in df.columns if c not in ("Date", "BO") and pd.api.types.is_numeric_dtype(df[c])]
    out = df[keep].dropna(subset=["Date"])
    out = out.groupby(["Date", "BO"], as_index=False).sum(numeric_only=True)
    target = DATA_DIR / "dashboard_data.xlsx"
    out.to_excel(target, index=False)
    print(f"Saved {len(out)} rows -> {target}")


if __name__ == "__main__":
    main()
