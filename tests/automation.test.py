#!/usr/bin/env python3
"""Offline tests for the CEPT MIS pipeline (no browser, no network).

    python3 tests/automation.test.py

They cover the two ways the pipeline can silently corrupt the dashboard's
per-BO numbers: offices that share a name (the master sheet has 82 such names)
must stay separate rows, and a run that downloaded nothing must not file the
previous run's figures under a new date.
"""
import datetime as dt
import json
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "automation"))
import cept_mis_cbs as cbs  # noqa: E402
import sanity_check as sanity  # noqa: E402

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

# supervisor: one bounded child process per SUB division, so a wedged MIS page costs
# the one sub division that is broken instead of a whole division (and with it the
# healthy sub divisions that shared its process)
divs = cbs.master_divisions()
check("master_divisions reads the division list offline (no browser)",
      len(divs) == 9 and "Ambala Division" in divs and "Sonipat Division" in divs, f"{divs}")

PAGE_DIVS = ["Ambala Division", "Bhiwani Division", "Faridabad Division", "Gurgaon Division",
             "Hissar Division", "Karnal Division", "Kurukshetra Division", "Rohtak Division",
             "Sonepat Division"]  # MIS spells it "Sonepat", the master sheet "Sonipat"
PAGE_SUBS = {
    "Ambala Division": ["Ambala Sub-Division", "Ambala City Sub- Division", "Panchkula Sub Division",
                        "Yamuna Nagar Sub Division"],
    "Bhiwani Division": ["Bhiwani Sub Division", "Mohindergarh Sub division", "Narnaul Sub Division"],
    "Faridabad Division": ["Faridabad West Sub Division", "FARIDABAD EAST SUB DN"],
    "Gurgaon Division": ["Gurgaon East Sub Division", "Rewari Sub Division"],
    "Hissar Division": ["Hissar Sub Division", "Sirsa-I Sub division"],
    "Karnal Division": ["Karnal East Sub Division"],
    "Kurukshetra Division": ["Kaithal Sub Division"],
    "Rohtak Division": ["Rohtak Sub Division", "Jhajjar Sub Division", "Bahadurgarh  sub division"],
    "Sonepat Division": ["Sonipat Sub division", "Jind East Sub Division"],
}


class FakeChild:
    """The lister records MIS's labels; Mohindergarh hangs on every attempt, Jhajjar
    fails its first attempt and recovers on the retry, everything else files a workbook.
    `date` is the report date the page resolves - a Sunday means "nothing to fetch"."""

    def __init__(self, date="2026-11-03"):
        self.seen, self.listed, self.attempts = [], 0, {}
        self.date = date

    def run(self, cmd, env, budget):
        if env.get("CEPT_LIST") == "1":
            self.listed += 1
            cbs.DIVISIONS_FILE.write_text(json.dumps({"date": self.date, "divisions": PAGE_DIVS,
                                                      "subs": PAGE_SUBS}), encoding="utf-8")
            return 0, 20.0
        div, sub = env["REPORT_DIVISIONS"], env.get("REPORT_SUBDIVISIONS")
        self.seen.append((div, sub, env.get("CEPT_CHILD"), env.get("SKIP_MERGE"), budget))
        key = (div, sub)
        self.attempts[key] = self.attempts.get(key, 0) + 1
        if sub == "Mohindergarh Sub division":
            cbs.record_window(self.date)   # a child dies mid-sweep, after resolving the date
            return None, float(budget)      # killed at the budget
        if sub == "Jhajjar Sub Division" and self.attempts[key] == 1:
            return 1, 5.0                   # transient: retried below
        cbs.record_window(self.date)
        if sub:
            cbs.part_file("accounts_opened", div, sub).write_bytes(b"not really an xlsx")
        return 0, 30.0


def patch_sweep(tmp, child, built, date=None):
    """Point every file the supervisor writes at a temp dir for one test block."""
    real = (cbs.run_bounded, cbs.build_outputs, cbs.EXPORT_DIR, cbs.WINDOW_FILE,
            cbs.DIVISIONS_FILE, cbs.COVERAGE_FILE)
    cbs.run_bounded = child.run
    cbs.build_outputs = lambda iso: built.append(iso)
    cbs.EXPORT_DIR = Path(tmp)
    cbs.WINDOW_FILE = Path(tmp) / "window.json"
    cbs.DIVISIONS_FILE = Path(tmp) / "divisions.json"
    cbs.COVERAGE_FILE = Path(tmp) / "coverage.json"
    return real


def restore_sweep(real):
    (cbs.run_bounded, cbs.build_outputs, cbs.EXPORT_DIR, cbs.WINDOW_FILE,
     cbs.DIVISIONS_FILE, cbs.COVERAGE_FILE) = real


with tempfile.TemporaryDirectory() as tmp:
    child = FakeChild()
    built = []
    real = patch_sweep(tmp, child, built)
    planned = sum(len(v) for v in PAGE_SUBS.values())
    try:
        cbs.supervise_divisions()
        first_pass = [s for s in child.seen if s[4] == cbs.SUB_BUDGET_S]
        check("supervisor asks the page for its own division + sub division labels",
              child.listed == 1, f"listed {child.listed}")
        check("one child per sub division, not per division",
              len(first_pass) == planned, f"{len(first_pass)} of {planned}")
        check("the sub division labels come from MIS, not the master sheet",
              ("Rohtak Division", "Bahadurgarh  sub division") in [(d, s) for d, s, *_ in child.seen],
              f"{[(d, s) for d, s, *_ in child.seen if d == 'Rohtak Division']}")
        check("every sub division runs as a download-only child",
              all(c == "1" and m == "1" for _, _, c, m, _ in child.seen), f"{child.seen[:2]}")
        check("a sub division that came back empty is retried until it reports",
              child.attempts[("Rohtak Division", "Jhajjar Sub Division")] == 2
              and cbs.got_data("Rohtak Division", "Jhajjar Sub Division"),
              f"{child.attempts[('Rohtak Division', 'Jhajjar Sub Division')]} attempts")
        check("a wedged sub division is retried, then named - never left silent",
              child.attempts[("Bhiwani Division", "Mohindergarh Sub division")] == 1 + cbs.SUB_RETRIES
              and not cbs.got_data("Bhiwani Division", "Mohindergarh Sub division"),
              f"{child.attempts[('Bhiwani Division', 'Mohindergarh Sub division')]} attempts")
        check("retry rounds get a longer budget than the first pass",
              all(b == cbs.RETRY_BUDGET_S for *_, b in child.seen if b != cbs.SUB_BUDGET_S),
              f"{[(n, b) for n, _, _, _, b in child.seen if b != cbs.SUB_BUDGET_S]}")
        cov = json.loads(cbs.COVERAGE_FILE.read_text(encoding="utf-8"))
        check("coverage.json names exactly which sub divisions are missing",
              cov["planned"] == planned and cov["filed"] == planned - 1
              and [m["subDivision"] for m in cov["missing"]] == ["Mohindergarh Sub division"]
              and cov["missing"][0]["exit"] == "killed", f"{cov}")
        check("a killed sub division costs only that sub division",
              cbs.got_data("Bhiwani Division", "Narnaul Sub Division")
              and cbs.got_data("Bhiwani Division", "Bhiwani Sub Division"), "neighbours survived")
        check("a division with most of its sub divisions covered is not called missing",
              cbs.division_has_data("Bhiwani Division")
              and not cbs.division_has_data("Nowhere Division"), "partial division")
        check("the day is still built when a sub division dies", built == ["2026-11-03"], f"{built}")
    finally:
        restore_sweep(real)


class DeadListChild(FakeChild):
    """The page visit that reads the division labels never gets anywhere."""

    def run(self, cmd, env, budget):
        if env.get("CEPT_LIST") == "1":
            return 1, 5.0
        return super().run(cmd, env, budget)


with tempfile.TemporaryDirectory() as tmp:
    child = DeadListChild()
    built = []
    real = patch_sweep(tmp, child, built)
    cbs.DIVISIONS_FILE = Path(tmp) / "never-written.json"
    try:
        cbs.supervise_divisions()
        whole = [s for s in child.seen if s[4] == cbs.DIVISION_BUDGET_S]
        check("a failed division-list visit falls back to the master sheet",
              "Sonipat Division" in [d for d, _, *_ in child.seen],
              f"{[d for d, _, *_ in child.seen][:2]}")
        check("with no sub division plan, one child sweeps a whole division",
              len(whole) == 9 and all(not s for _, s, *_ in child.seen),
              f"{[(d, s, b) for d, s, _, _, b in whole][:2]}")
        cov = json.loads(cbs.COVERAGE_FILE.read_text(encoding="utf-8"))
        check("a day with no coverage at all is reported, not hidden",
              cov["planned"] == 9 and cov["filed"] == 0 and built == ["2026-11-03"],
              f"{cov['filed']}/{cov['planned']}, built {built}")
    finally:
        restore_sweep(real)


with tempfile.TemporaryDirectory() as tmp:
    child = FakeChild(date="2026-11-01")  # a Sunday: every CEPT office is closed
    built = []
    real = patch_sweep(tmp, child, built)
    try:
        cbs.supervise_divisions()
        check("no report is fetched for a Sunday, and no day is committed",
              child.listed == 1 and child.seen == [] and built == []
              and not cbs.COVERAGE_FILE.exists(),
              f"{len(child.seen)} children, built {built}")
    finally:
        restore_sweep(real)


class MondayToday(dt.date):
    """A Monday: the default window would land on Sunday if left alone."""

    @classmethod
    def today(cls):
        return cls(2026, 11, 9)


env_names = ("REPORT_START_DATE", "REPORT_END_DATE")
real_env = {k: os.environ.get(k) for k in env_names}
real_date = cbs.date
cbs.date = MondayToday
try:
    start, end = cbs.resolve_window(None)
    check("Monday's default window rolls back to Saturday instead of fetching a Sunday",
          (start.isoformat(), end.isoformat()) == ("2026-11-07", "2026-11-07"),
          f"{start} -> {end}")
    for k in env_names:
        os.environ[k] = "08-11-2026"
    start, end = cbs.resolve_window(None)
    check("an operator-asked Sunday stays a Sunday so the sweep can refuse it",
          cbs.holiday_window(start, end) and start.isoformat() == "2026-11-08",
          f"{start} -> {end}")
    os.environ["REPORT_START_DATE"] = "07-11-2026"
    os.environ["REPORT_END_DATE"] = "08-11-2026"
    start, end = cbs.resolve_window(None)
    check("a window that spans a working day is fetched even if it ends on a Sunday",
          (start.isoformat(), end.isoformat()) == ("2026-11-07", "2026-11-08")
          and not cbs.holiday_window(start, end), f"{start} -> {end}")
    check("is_sunday_text reads the ISO dates the pipeline actually carries",
          cbs.is_sunday_text("2026-11-01") and not cbs.is_sunday_text("2026-11-02")
          and not cbs.is_sunday_text(None), "sunday text")
finally:
    cbs.date = real_date
    for k, v in real_env.items():
        if v is None:
            os.environ.pop(k, None)
        else:
            os.environ[k] = v

FULL_COVERAGE = {"date": "2026-11-03", "planned": 31, "filed": 31, "missing": []}
ONE_GAP = {"date": "2026-11-03", "planned": 31, "filed": 27, "missing": [
    {"division": "Bhiwani Division", "subDivision": "Mohindergarh Sub division", "exit": "killed"}]}
check("a fully covered day passes the gate even when few BOs opened accounts",
      sanity.coverage_failures(FULL_COVERAGE) == [], f"{sanity.coverage_failures(FULL_COVERAGE)}")
check("a day that is missing a sub division is named and blocked",
      len(sanity.coverage_failures(ONE_GAP)) == 1
      and "Mohindergarh" in sanity.coverage_failures(ONE_GAP)[0]
      and "(killed)" in sanity.coverage_failures(ONE_GAP)[0],
      f"{sanity.coverage_failures(ONE_GAP)}")
check("SANITY_ALLOW_MISSING_SUBDIVISIONS lets an operator ship a known gap",
      sanity.coverage_failures(ONE_GAP, allow=1) == [], "allow=1")

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
