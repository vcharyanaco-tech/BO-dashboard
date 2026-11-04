#!/usr/bin/env python3
"""Offline tests for the CEPT MIS pipeline (no browser, no network).

    python3 tests/automation.test.py

They cover the two ways the pipeline can silently corrupt the dashboard's
per-BO numbers: offices that share a name (the master sheet has 82 such names)
must stay separate rows, and a run that downloaded nothing must not file the
previous run's figures under a new date.
"""
import json
import os
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
    write_book(export)
    return export


def write_book(export, bump=False):
    """The synthetic report, optionally with one figure moved (a real next day)."""
    rows = [list(r) for r in ROWS]
    if bump:
        rows[0][3] += 1  # 'Balu B.O' J6645 MIS 10 -> 11
    pd.DataFrame([HEADER] + rows).to_excel(export / "accounts_opened_merged.xlsx", index=False, header=False)


with tempfile.TemporaryDirectory() as tmp:
    export = sandbox(tmp)

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

    # a second day with different figures adds its own rows for the same two offices
    write_book(export, bump=True)
    cbs.append_history("2026-11-04")
    hist3 = pd.read_csv(cbs.HISTORY_FILE, dtype=str)
    check("each day keeps both twins", len(hist3) == 8 and
          len(hist3[hist3["bo"] == "Balu B.O"]) == 4, f"{len(hist3)} rows")

    # same report filed under a new date: one snapshot, one day
    cbs.append_history("2026-11-05")
    hist4 = pd.read_csv(cbs.HISTORY_FILE, dtype=str)
    days = sorted(hist4["date"].unique())
    check("an identical snapshot is not filed as an extra day",
          days == ["2026-11-03", "2026-11-05"] and len(hist4) == 8, f"days {days}, {len(hist4)} rows")
    check("the duplicate keeps the newer stamp",
          hist4[hist4["date"] == "2026-11-05"].shape[0] == 4 and "2026-11-04" not in days,
          f"{days}")

    # row_identity: code wins, name only as a fallback, office rows namespaced
    check("row_identity prefers the code", cbs.row_identity("bo", "Balu B.O", "j6645") == "J6645")
    check("row_identity falls back to the canonical name",
          cbs.row_identity("bo", "Balu bo", "") == cbs.norm_bo_name("Balu B.O"))
    check("office rows never collide with BO rows", cbs.row_identity("so", "Balu B.O", "J6645").startswith("so:"))

# supervisor: one child process per division, each with a hard time budget, so a
# wedged MIS page costs one division instead of the whole day
divs = cbs.master_divisions()
check("master_divisions reads the division list offline (no browser)",
      len(divs) == 9 and "Ambala Division" in divs and "Sonipat Division" in divs, f"{divs}")

PAGE_DIVS = ["Ambala Division", "Bhiwani Division", "Faridabad Division", "Gurgaon Division",
             "Hissar Division", "Karnal Division", "Kurukshetra Division", "Rohtak Division",
             "Sonepat Division"]  # MIS spells it "Sonepat", the master sheet "Sonipat"


class FakeChild:
    """Children stand-in: the lister records labels, Bhiwani hangs, Rohtak exits badly."""

    def __init__(self):
        self.seen, self.listed = [], 0

    def run(self, cmd, env, budget):
        if env.get("CEPT_LIST") == "1":
            self.listed += 1
            cbs.DIVISIONS_FILE.write_text(json.dumps({"date": "2026-11-03",
                                                      "divisions": PAGE_DIVS}), encoding="utf-8")
            return 0, 20.0
        div = env["REPORT_DIVISIONS"]
        self.seen.append((div, env.get("CEPT_CHILD"), env.get("SKIP_MERGE")))
        if div == "Bhiwani Division":
            cbs.record_window("2026-11-03")  # a child dies mid-sweep, after resolving
            return None, float(budget)       # killed at the budget
        if div == "Rohtak Division":
            return 1, 12.0
        cbs.record_window("2026-11-03")
        return 0, 30.0


with tempfile.TemporaryDirectory() as tmp:
    child = FakeChild()
    built = []
    real = (cbs.run_bounded, cbs.build_outputs, cbs.part_files, cbs.EXPORT_DIR,
            cbs.WINDOW_FILE, cbs.DIVISIONS_FILE)
    cbs.run_bounded = child.run
    cbs.build_outputs = lambda iso: built.append(iso)
    cbs.part_files = lambda: []
    cbs.EXPORT_DIR = Path(tmp)
    cbs.WINDOW_FILE = Path(tmp) / "window.json"
    cbs.DIVISIONS_FILE = Path(tmp) / "divisions.json"
    try:
        check("supervisor asks the page for its own division labels",
              cbs.supervise_divisions() and child.listed == 1, f"listed {child.listed}")
        check("supervisor sweeps every division", len(child.seen) == 9, f"{len(child.seen)} children")
        check("the division names come from MIS, not the master sheet",
              "Sonepat Division" in [d for d, _, _ in child.seen], f"{child.seen[:2]}")
        check("every division runs as a download-only child",
              all(mode == ("1", "1") for _, mode in [(d, (c, m)) for d, c, m in child.seen]),
              f"{child.seen[:2]}")
        check("a killed division still lets the day be built", built == ["2026-11-03"], f"{built}")
    finally:
        (cbs.run_bounded, cbs.build_outputs, cbs.part_files, cbs.EXPORT_DIR,
         cbs.WINDOW_FILE, cbs.DIVISIONS_FILE) = real


class DeadListChild(FakeChild):
    """The page visit that reads the division labels never gets anywhere."""

    def run(self, cmd, env, budget):
        if env.get("CEPT_LIST") == "1":
            return 1, 5.0
        return super().run(cmd, env, budget)


with tempfile.TemporaryDirectory() as tmp:
    child = DeadListChild()
    real = (cbs.run_bounded, cbs.build_outputs, cbs.part_files, cbs.EXPORT_DIR,
            cbs.WINDOW_FILE, cbs.DIVISIONS_FILE)
    cbs.run_bounded = child.run
    cbs.build_outputs = lambda iso: None
    cbs.part_files = lambda: []
    cbs.EXPORT_DIR = Path(tmp)
    cbs.WINDOW_FILE = Path(tmp) / "window.json"
    cbs.DIVISIONS_FILE = Path(tmp) / "never-written.json"
    try:
        check("a failed division-list visit falls back to the master sheet",
              cbs.supervise_divisions() and "Sonipat Division" in [d for d, _, _ in child.seen],
              f"{[d for d, _, _ in child.seen]}")
    finally:
        (cbs.run_bounded, cbs.build_outputs, cbs.part_files, cbs.EXPORT_DIR,
         cbs.WINDOW_FILE, cbs.DIVISIONS_FILE) = real

rc, secs = cbs.run_bounded([sys.executable, "-c", "import time; time.sleep(60)"], dict(os.environ), 2)
check("run_bounded kills a child that overspends its budget", rc is None and secs < 15,
      f"rc {rc}, {secs:.1f}s")
rc, secs = cbs.run_bounded([sys.executable, "-c", "raise SystemExit(3)"], dict(os.environ), 30)
check("run_bounded hands back the child's exit code", rc == 3, f"rc {rc}")

# select_radio: Submit's postback is often still in flight when the radio is dispatched,
# and an evaluate that lands mid-navigation used to end the whole sub division - the 03/11
# run lost 27 of 30 sub divisions that way while the debug page showed the report rendered.
DESTROYED = "Page.evaluate: Execution context was destroyed, most likely because of a navigation"


class NavPage:
    """A page that fails the first `fails` evaluations, then takes the click."""

    def __init__(self, fails):
        self.fails, self.clicks = fails, 0

    def evaluate(self, js, sel):
        if self.fails:
            self.fails -= 1
            raise RuntimeError(DESTROYED)
        self.clicks += 1

    def wait_for_load_state(self, state, timeout=None):
        return None


class DeadPage(NavPage):
    """A page that is never not navigating."""

    def __init__(self):
        super().__init__(0)

    def evaluate(self, js, sel):
        raise RuntimeError(DESTROYED)


real_sleep = cbs.time.sleep
cbs.time.sleep = lambda s: None  # keep the settle loop instant in the test
try:
    page = NavPage(2)
    cbs.select_radio(page, "#accounts_opened")
    check("a navigation mid-select_radio is retried, not fatal", page.clicks == 1, f"clicks {page.clicks}")

    page = NavPage(0)
    cbs.select_radio(page, "#accounts_opened")
    check("a clean page is still dispatched once", page.clicks == 1, f"clicks {page.clicks}")

    try:
        cbs.select_radio(DeadPage(), "#accounts_opened", settle_ms=0)
        check("a page that never settles still raises", False, "no error raised")
    except RuntimeError as e:
        check("a page that never settles still raises", str(e) == DESTROYED, str(e))
finally:
    cbs.time.sleep = real_sleep

print("ALL PASS" if failures == 0 else f"FAILED: {failures} check(s)")
sys.exit(1 if failures else 0)
