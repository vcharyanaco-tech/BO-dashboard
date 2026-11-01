"""Smoke test for the deployed dashboard.

Loads the live page in headless Chromium and asserts that the auto-fed data
actually rendered: KPI cards exist, the BO table has rows, the auto-load status
is not the "not available" fallback, and no uncaught page errors occurred.

Exit 0 = healthy, 1 = broken. Override the target with DASHBOARD_URL, e.g.:
    DASHBOARD_URL=http://localhost:8000/ python -u automation/smoke_test.py
"""
import os
import sys

from playwright.sync_api import sync_playwright

URL = os.getenv("DASHBOARD_URL", "https://vcharyanaco-tech.github.io/BO-dashboard/")


def main():
    errors = []
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
            print(f"FAIL: no KPI cards rendered within 45s. autoStatus: {status!r}")
            browser.close()
            return 1
        kpis = page.locator("#kpis .kpi").count()
        rows = page.locator("#tbl tbody tr").count()
        status = page.locator("#autoStatus").inner_text()
        asof = page.locator("#dataAsOf").inner_text()
        print(f"smoke: KPI cards={kpis}, table rows={rows}, data as of={asof}")
        print(f"smoke: autoStatus: {status}")
        ok = kpis > 0 and rows > 0 and not errors
        if "available nahi" in status:
            print("smoke: FAIL - auto data did not load (fallback message shown)")
            ok = False
        if errors:
            print("smoke: page errors:")
            for e in errors:
                print("  " + e)
            ok = False
        browser.close()
        print("smoke: PASS" if ok else "smoke: FAIL")
        return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
