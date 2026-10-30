"""Backfill: build data/dashboard_data.xlsx from the 3 existing merged KPI files."""
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
EXPORT_DIR = ROOT / "data" / "exports"
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

METRICS = [
    ("accounts_opened_merged.xlsx", "Accounts Opened"),
    ("accounts_closed_merged.xlsx", "Accounts Closed"),
    ("net_addition_merged.xlsx", "Net Addition"),
]


def log(msg):
    print(msg, flush=True)


def main():
    yesterday = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    combined = None
    for fname, label in METRICS:
        f = EXPORT_DIR / fname
        if not f.exists():
            log(f"{fname}: missing, skipping")
            continue
        df = pd.read_excel(f)
        bo_col = next((c for c in df.columns if str(c).strip().lower() in ("name", "bo", "bo name", "branch office")), None)
        if not bo_col:
            log(f"{fname}: no BO name column found, columns: {list(df.columns)}")
            continue
        num_cols = [c for c in df.columns if c != bo_col and pd.api.types.is_numeric_dtype(df[c])]
        if not num_cols:
            log(f"{fname}: no numeric columns found")
            continue
        df[label] = df[num_cols].sum(axis=1)
        part = df[[bo_col, label]].rename(columns={bo_col: "BO"}).groupby("BO", as_index=False)[label].sum()
        combined = part if combined is None else combined.merge(part, on="BO", how="outer")
        log(f"{fname}: {len(part)} BO rows, metric {label}")
    if combined is None:
        raise SystemExit("No merged KPI files found to build dashboard data.")
    combined.insert(0, "Date", yesterday)
    for _, label in METRICS:
        if label not in combined.columns:
            combined[label] = 0
    combined = combined.fillna(0)
    out = DATA_DIR / "dashboard_data.xlsx"
    combined.to_excel(out, index=False)
    log(f"Dashboard data written: {out} ({len(combined)} BO rows)")


if __name__ == "__main__":
    main()
