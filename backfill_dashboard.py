"""Backfill: build data/dashboard_data.xlsx from the 3 existing merged KPI files.

Merged files have title rows at top and duplicate column names;
 detect the header row containing 'Name' or 'SOL ID / BOCODE',
 dedupe column names, and parse data from there using positional indexing.
"""
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


def find_header_row(matrix):
    for i, row in enumerate(matrix):
        cells = [str(c).strip() if c is not None else "" for c in row]
        if any(c.lower() in ("name", "bo", "bo name", "branch office") or "sol id" in c.lower() or "bocode" in c.lower() for c in cells):
            return i
    return -1


def dedupe(cols):
    seen = {}
    out = []
    for c in cols:
        c = str(c).strip() if c else "col"
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out


def main():
    yesterday = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    combined = None
    for fname, label in METRICS:
        f = EXPORT_DIR / fname
        if not f.exists():
            log(f"{fname}: missing, skipping")
            continue
        matrix = pd.read_excel(f, header=None).values.tolist()
        hIdx = find_header_row(matrix)
        if hIdx < 0:
            log(f"{fname}: header row not found, skipping")
            continue
        raw_headers = [str(c).strip() if c is not None else "" for c in matrix[hIdx]]
        headers = dedupe(raw_headers)
        rows = [r for r in matrix[hIdx+1:] if any(str(c).strip() for c in r if c is not None)]
        df = pd.DataFrame(rows, columns=headers)
        # Find BO column positionally
        bo_pos = next((i for i, h in enumerate(headers) if h.lower() in ("name", "bo", "bo name", "branch office") or "sol id" in h.lower() or "bocode" in h.lower()), None)
        if bo_pos is None:
            log(f"{fname}: no BO column found, headers: {headers}")
            continue
        bo_series = df.iloc[:, bo_pos].astype(str).str.strip()
        # Numeric coercion positionally
        num_pos = []
        for i in range(len(headers)):
            if i == bo_pos:
                continue
            col = df.iloc[:, i]
            coerced = pd.to_numeric(col, errors="coerce")
            if coerced.notna().any():
                num_pos.append(i)
        if not num_pos:
            log(f"{fname}: no numeric columns found")
            continue
        metric_vals = df.iloc[:, num_pos].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
        part = pd.DataFrame({"BO": bo_series, label: metric_vals}).groupby("BO", as_index=False)[label].sum()
        combined = part if combined is None else combined.merge(part, on="BO", how="outer")
        log(f"{fname}: header at row {hIdx}, {len(part)} BO rows, metric {label}")
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
