"""Sanity-check freshly merged KPI exports before CI commits them.

For every data/exports/{kpi}_merged.xlsx in the worktree, compare its BO data-row
count with the copy committed at HEAD. Exit 1 if a new file is empty or has lost
more than half of its BO rows - that signals a partial/failed scrape, and the old
(good) file should be kept instead of being overwritten.

Row counts cannot catch a *copy*: if the worktree file is byte-identical to HEAD,
the run served/fetched the same report again and the new date would be a duplicate
snapshot, so that is printed as a warning.

Run after automation/cept_mis_cbs.py and before the commit step:
    python -u automation/sanity_check.py
"""
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent.parent
EXPORT_DIR = ROOT / "data" / "exports"
MAX_DROP_FRACTION = 0.5  # fail if BO rows drop by more than 50% vs HEAD

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


def main():
    merged = sorted(EXPORT_DIR.glob("*_merged.xlsx"))
    if not merged:
        print("sanity: no *_merged.xlsx files found - nothing to check")
        return 0
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
            failures.append(
                f"{rel}: BO rows dropped {old_n} -> {new_n} "
                f"(more than {int(MAX_DROP_FRACTION * 100)}% loss)"
            )
    if failures:
        print("sanity: FAIL - not committing suspicious exports:")
        for f in failures:
            print("  " + f)
        return 1
    print("sanity: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
