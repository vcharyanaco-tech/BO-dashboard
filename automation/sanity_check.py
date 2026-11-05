"""Sanity-check freshly merged KPI exports before CI commits them.

What makes a day trustworthy is *coverage*, not volume. The MIS sweep downloads one
workbook per sub division (the plan lives in debug/coverage.json), so a day is complete
when every planned sub division reported - and a day with only a handful of BOs opening
accounts is perfectly real data (a quiet day, or a week that starts slowly). Only when
there is no coverage report (a run outside the supervisor) does the old row-count rule
apply, because then a large drop really is the only signal available.

So:
  * 0 BO rows in a merged file               -> fail, never commit that.
  * planned sub divisions that never filed   -> fail, and name them. Override with
    SANITY_ALLOW_MISSING_SUBDIVISIONS=<n> (the dispatch input
    "allow_missing_subdivisions") if MIS is known to be wedged and the owner wants the
    partial day anyway.
  * a sub division whose report came back with zero rows -> covered, not missing. A quiet
    day at one sub division is real data, so it never blocks the commit.
  * a big row drop vs HEAD                   -> warn when coverage says the day is whole.
  * byte-identical to HEAD                   -> warn: the site served the same snapshot
    again, and history's "one snapshot = one day" rule will not file it twice.

Run after automation/cept_mis_cbs.py and before the commit step:
    python -u automation/sanity_check.py
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent.parent
EXPORT_DIR = ROOT / "data" / "exports"
COVERAGE_FILE = ROOT / "debug" / "coverage.json"
MAX_DROP_FRACTION = 0.5  # warn (or fail without a coverage report) on a drop this big

HEADER_HINTS = ("name", "bo", "bo name", "branch office")


def cell_text(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none") else t


def is_header(cells):
    return any(
        c.lower() in HEADER_HINTS or "sol id" in c.lower() or "bocode" in c.lower()
        for c in cells
    )


def bo_row_count(path):
    """Count real BO data rows (header/Total/title rows excluded)."""
    try:
        matrix = pd.read_excel(path, header=None).values.tolist()
    except Exception as e:
        print(f"  cannot read {path}: {e}")
        return 0
    h = -1
    for i, row in enumerate(matrix):
        if is_header([cell_text(c) for c in row]):
            h = i
            break
    if h < 0:
        return 0
    n = 0
    for row in matrix[h + 1:]:
        cells = [cell_text(c) for c in row]
        if not any(cells):
            continue
        if cells[0].lower().startswith("total") or is_header(cells):
            continue
        n += 1
    return n


def head_blob(rel_path):
    """Bytes of the file as committed at HEAD (None if never committed)."""
    try:
        return subprocess.run(
            ["git", "show", f"HEAD:{rel_path}"], capture_output=True, check=True
        ).stdout
    except Exception:
        return None


def head_bo_row_count(rel_path):
    blob = head_blob(rel_path)
    if blob is None:
        return None
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as t:
            t.write(blob)
            tmp = t.name
        return bo_row_count(Path(tmp))
    except Exception:
        return None
    finally:
        if tmp:
            os.unlink(tmp)


def load_coverage():
    """The supervisor's report on which sub divisions downloaded; None if it never ran."""
    try:
        return json.loads(COVERAGE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def allowed_missing():
    """How many sub divisions the operator agreed to ship without.

    Set from the workflow dispatch input, so a blank or nonsense value simply keeps the
    gate closed rather than crashing the run.
    """
    try:
        return max(0, int(os.getenv("SANITY_ALLOW_MISSING_SUBDIVISIONS", "").strip() or 0))
    except ValueError:
        return 0


def coverage_failures(coverage, allow=0):
    """Name every planned sub division that did not report, unless we allow that many.

    A day that is missing a sub division is missing whole offices, so its per-BO rows and
    its sum are both wrong - unlike a quiet day, where the offices that are present are
    all the offices that actually opened accounts.
    """
    missing = coverage.get("missing") or []
    if len(missing) <= allow:
        return []
    listed = ", ".join(f"{m.get('division')} / {m.get('subDivision')} "
                       f"({m.get('exit', 'unknown')})" for m in missing)
    return [f"{len(missing)} of {coverage.get('planned')} sub divisions never downloaded "
            f"after retries: {listed}"]


def main():
    merged = sorted(EXPORT_DIR.glob("*_merged.xlsx"))
    if not merged:
        print("sanity: no *_merged.xlsx files found - nothing to check")
        return 0
    coverage = load_coverage()
    allow = allowed_missing()
    if coverage:
        quiet = len(coverage.get("empty") or [])
        extra = f", {quiet} answered with 0 rows" if quiet else ""
        print(f"sanity: coverage report for {coverage.get('date')}: "
              f"{coverage.get('filed')}/{coverage.get('planned')} sub divisions filed{extra}")
        if allow:
            print(f"sanity: SANITY_ALLOW_MISSING_SUBDIVISIONS={allow} - up to {allow} "
                  f"sub division(s) may be missing and the day still commits")
    else:
        print("sanity: no debug/coverage.json - this merge did not come from the "
              "sub-division supervisor, so the row-count gate applies")
    failures = []
    for f in merged:
        rel = f.relative_to(ROOT).as_posix()
        new_n = bo_row_count(f)
        old_n = head_bo_row_count(rel)
        old_txt = old_n if old_n is not None else "n/a (new file)"
        print(f"sanity: {rel}: worktree={new_n} BO rows, HEAD={old_txt}")
        blob = head_blob(rel)
        if blob is not None and blob == f.read_bytes():
            # Row counts alone cannot see this: the site served the same report
            # again (or the scrape downloaded nothing), and the day would be filed
            # as a fresh snapshot that is really a copy of the previous one.
            print(f"WARN sanity: {rel} is byte-identical to HEAD - this run fetched "
                  f"the same report again, so the new date is a duplicate snapshot")
        if new_n == 0:
            failures.append(f"{rel}: 0 BO rows in worktree")
        elif old_n and new_n < old_n * (1 - MAX_DROP_FRACTION):
            drop = f"{rel}: BO rows {old_n} -> {new_n} ({int(new_n / old_n * 100)}% of yesterday)"
            if coverage:
                # Coverage, not volume, decides - few BOs opening accounts on a given
                # day is ordinary, and Sundays are not fetched at all.
                print(f"WARN sanity: {drop} - volume is not the gate, the coverage "
                      f"report below decides")
            else:
                failures.append(f"{drop}; more than {int(MAX_DROP_FRACTION * 100)}% loss "
                                f"and no coverage report to tell a quiet day from a bad scrape")
    if coverage:
        failures += coverage_failures(coverage, allow)
    if failures:
        print("sanity: FAIL - not committing suspicious exports:")
        for f in failures:
            print("  " + f)
        return 1
    print("sanity: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
