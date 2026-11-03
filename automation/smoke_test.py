"""Smoke test for the deployed dashboard.

Loads the live page in headless Chromium and asserts that the auto-fed data
rendered AND that the interactive controls actually work (the CSP forbids
inline event handlers, so controls must be wired via addEventListener in
app.js — this test fires them for real):

  1. KPI cards + all BO-master rows in the table render, no page/console errors
  2. Quick view = Top 10 applies ("top 10" in the info line, <= 10 rows)
  3. Group by = SO-wise switches the table title and shrinks the row count
  4. BO text filter narrows the table

Exit 0 = healthy, 1 = broken. Override the target with DASHBOARD_URL, e.g.:
    DASHBOARD_URL=http://localhost:8000/ python -u automation/smoke_test.py
"""
import os
import re
import sys

from playwright.sync_api import sync_playwright

URL = os.getenv("DASHBOARD_URL", "https://vcharyanaco-tech.github.io/BO-dashboard/")


def main():
    errors = []
    failures = []

    def check(name, cond, detail=""):
        print(f"smoke: {'PASS' if cond else 'FAIL'} - {name}" + (f" ({detail})" if detail else ""))
        if not cond:
            failures.append(name)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        page = browser.new_page()
        page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
        page.on(
            "console",
            lambda m: errors.append(f"console.error: {m.text}") if m.type == "error" else None,
        )
        print(f"smoke: loading {URL}")
        page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_selector("#kpis .kpi", timeout=45000)
        except Exception:
            status = ""
            try:
                status = page.locator("#autoStatus").inner_text()
            except Exception:
                pass
            print(f"smoke: FAIL - no KPI cards rendered within 45s. autoStatus: {status!r}")
            browser.close()
            return 1

        kpis = page.locator("#kpis .kpi").count()
        rows_all = page.locator("#tbl tbody tr").count()
        master_count = page.evaluate("() => (typeof BO_MASTER !== 'undefined') ? BO_MASTER.length : 0")
        status = page.locator("#autoStatus").inner_text()
        asof = page.locator("#dataAsOf").inner_text()
        print(f"smoke: KPI cards={kpis}, table rows={rows_all}, master BOs={master_count}, data as of={asof}")
        print(f"smoke: autoStatus: {status}")
        check("KPI cards rendered", kpis > 0)
        check("BO table has rows", rows_all > 0, f"{rows_all} rows")
        check("all BO master rows rendered", master_count > 0 and rows_all >= master_count,
              f"{rows_all} rows for {master_count} master BOs")
        check("auto data loaded", "available nahi" not in status, status[:80])
        check("data-as-of uses DD-MM-YYYY", bool(re.fullmatch(r"\d{2}-\d{2}-\d{4}", asof)), asof)
        labels = page.evaluate(
            "() => { const c = Chart.getChart(document.getElementById('chart')); return c ? c.data.labels : []; }"
        )
        check("chart labels use DD-MM-YYYY",
              bool(labels) and all(re.fullmatch(r"\d{2}-\d{2}-\d{4}", str(v)) for v in labels),
              f"{labels[:3]}")

        # --- interaction 1: quick view = Top 10 (select fires 'change') ---
        page.select_option("#quick", "top")
        page.wait_for_timeout(400)
        info = page.locator("#tblInfo").inner_text()
        rows_top = page.locator("#tbl tbody tr").count()
        check("quick view Top 10 applies", "top 10" in info and rows_top <= 10,
              f"tblInfo={info!r}, rows={rows_top}")
        page.select_option("#quick", "all")
        page.wait_for_timeout(300)

        # --- interaction 2: group by SO (select fires 'change') ---
        page.select_option("#groupBy", "so")
        page.wait_for_timeout(400)
        title = page.locator("#tblTitle").inner_text()
        rows_so = page.locator("#tbl tbody tr").count()
        check("group by SO-wise applies", "SO-wise" in title and 0 < rows_so < rows_all,
              f"title={title!r}, rows={rows_so} vs {rows_all}")
        head_txt = page.locator("#tbl thead").inner_text()
        check("SO view keeps ΣBO and SO-reported columns separate",
              "ΣBO" in head_txt and "(SO)" in head_txt, head_txt[:100])
        page.select_option("#groupBy", "bo")
        page.wait_for_timeout(300)

        # --- interaction 3: BO text filter (fires 'input', debounced 150ms) ---
        page.fill("#bo", "a")
        page.wait_for_timeout(600)
        rows_f = page.locator("#tbl tbody tr").count()
        check("BO filter narrows table", 0 <= rows_f < rows_all, f"rows={rows_f} vs {rows_all}")
        page.fill("#bo", "")
        page.wait_for_timeout(600)

        if errors:
            print("smoke: page errors:")
            for e in errors:
                print("  " + e)
        check("no page/console errors", not errors, f"{len(errors)} error(s)")

        browser.close()
        if failures:
            print(f"smoke: FAIL - {len(failures)} check(s) failed: {failures}")
            return 1
        print("smoke: PASS")
        return 0


if __name__ == "__main__":
    sys.exit(main())
