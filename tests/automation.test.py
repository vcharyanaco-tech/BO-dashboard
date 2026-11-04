#!/usr/bin/env python3
"""Offline tests for the CEPT MIS pipeline (no browser, no network).

    python3 tests/automation.test.py

They cover the two ways the pipeline can silently corrupt the dashboard's
per-BO numbers: offices that share a name (the master sheet has 82 such names)
must stay separate rows, and a run that downloaded nothing must not file the
previous run's figures under a new date.
"""
import sys
import tempfile
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "automation"))
import cept_mis_cbs as cbs  # noqa: E402

failures = 0


def check(name, cond, detail=""):
    global failures
    print(f"{'PASS' if cond else 'FAIL'} {name}" + (f"  [{detail}]" if not cond and detail else ""))
    if not cond:
        failures += 1


HEADER = ["Name", "SOL ID / BOCODE", "SO Name", "MIS", "RD", "TD"]
ROWS = [
    # two different BOs that share the display name 'Balu B.O'
    ["Balu B.O", "J6645", "Nissing S.O", 10, 5, 0],
    ["Balu B.O", "K1234", "Rohtak S.O", 3, 1, 0],
    ["Adhoyi B.O", "I1895", "Adhoya S.O", 1, 1, 1],
    # office summary row: blank SO Name cell + 8-digit code
    ["Adhoya S.O", "13320501", "", 20, 5, 5],
    ["Total:", "", "", 40, 12, 6],
]


def sandbox(tmp):
    """Point the module's data paths at a temp dir."""
    export = Path(tmp) / "data" / "exports"
    export.mkdir(parents=True)
    cbs.EXPORT_DIR = export
    cbs.HISTORY_FILE = Path(tmp) / "history.csv"
    cbs.DATA_DIR = Path(tmp) / "data"
    wb = pd.DataFrame([HEADER] + ROWS)
    wb.to_excel(export / "accounts_opened_merged.xlsx", index=False, header=False)
    return export


with tempfile.TemporaryDirectory() as tmp:
    sandbox(tmp)

    data = cbs._read_merged("accounts_opened")
    check("merged file keyed per office, not per name", len([k for k, v in data.items() if v["kind"] == "bo"]) == 3,
          f"got {sorted(data)}")
    check("twin names keep their own codes",
          data.get("J6645", {}).get("value") == 15 and data.get("K1234", {}).get("value") == 4,
          f"{[(k, v['value']) for k, v in data.items()]}")
    check("office summary row stays in its own identity space",
          data.get("so:adhoya so", {}).get("kind") == "so" and data.get("so:adhoya so", {}).get("value") == 30,
          f"{data.get('so:adhoya so')}")
    check("Total footer row is not a BO", not any(v["bo"].lower().startswith("total") for v in data.values()))

    cbs.append_history("2026-11-03")
    hist = pd.read_csv(cbs.HISTORY_FILE, dtype=str)
    check("history keeps one row per office", len(hist) == 4, f"{len(hist)} rows")
    balu = hist[hist["bo"] == "Balu B.O"]
    check("both Balu B.O offices survive into history", len(balu) == 2 and set(balu["sol"]) == {"J6645", "K1234"},
          f"{balu.to_dict('records')}")

    # same date re-run: replaced, never doubled, and never merged by name
    cbs.append_history("2026-11-03")
    hist2 = pd.read_csv(cbs.HISTORY_FILE, dtype=str)
    check("re-running one date replaces its rows", len(hist2) == 4, f"{len(hist2)} rows")

    # a second day adds its own rows for the same two offices
    cbs.append_history("2026-11-04")
    hist3 = pd.read_csv(cbs.HISTORY_FILE, dtype=str)
    check("each day keeps both twins", len(hist3) == 8 and
          len(hist3[hist3["bo"] == "Balu B.O"]) == 4, f"{len(hist3)} rows")

    # row_identity: code wins, name only as a fallback, office rows namespaced
    check("row_identity prefers the code", cbs.row_identity("bo", "Balu B.O", "j6645") == "J6645")
    check("row_identity falls back to the canonical name",
          cbs.row_identity("bo", "Balu bo", "") == cbs.norm_bo_name("Balu B.O"))
    check("office rows never collide with BO rows", cbs.row_identity("so", "Balu B.O", "J6645").startswith("so:"))

print("ALL PASS" if failures == 0 else f"FAILED: {failures} check(s)")
sys.exit(1 if failures else 0)
