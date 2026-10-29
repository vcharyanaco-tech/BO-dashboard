"""
CEPT MIS CBSTrans3 BO-wise automation.

Cascading dropdowns wait for specific option text; submit button detected
flexibly (input/button/link) since its exact tag may vary.
"""
import io, sys, traceback
from datetime import date
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
DEBUG_DIR = ROOT / "debug"
DATA_DIR.mkdir(exist_ok=True)
DEBUG_DIR.mkdir(exist_ok=True)

REPORT_URL = "https://mis.cept.gov.in/CBS/CBSTrans3.aspx"
CIRCLE = "Haryana Circle"
REGION = "DivReporting Circle"
METRICS = [
    ("#ContentPlaceHolder1_ACO", "Accounts Opened"),
    ("#ContentPlaceHolder1_ACC", "Accounts Closed"),
    ("#ContentPlaceHolder1_NET", "Net Addition"),
]

CIRCLE_SEL = "#ContentPlaceHolder1_CircleDropDown"
REGION_SEL = "#ContentPlaceHolder1_RegionDropDown"
DIV_SEL = "#ContentPlaceHolder1_DivisionDropDown"
SUB_SEL = "#ContentPlaceHolder1_SubDivisionDropDown"

SUBMIT_SELECTORS = [
    "input[type=submit]",
    "button[type=submit]",
    "a:has-text('Show')",
    "a:has-text('Report')",
    "a:has-text('Generate')",
    "button:has-text('Show')",
    "button:has-text('Report')",
    "[onclick*='Show']",
    "[onclick*='Report']",
    "[onclick*='Generate']",
]


def norm(s):
    return "".join(str(s).lower().split())


def save_debug(page, tag):
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{tag}.png"), full_page=True)
        (DEBUG_DIR / f"{tag}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass


def open_page(page):
    page.goto(REPORT_URL, wait_until="networkidle")
    date_input = page.query_selector("input[id*='Date'], input[name*='Date'], input[type='text']")
    if date_input:
        date_input.fill(date.today().strftime("%d/%m/%Y"))


def wait_for_option(page, sel, text, timeout=60000):
    t = norm(text)
    page.wait_for_function(
        "([sel, t]) => { const s = document.querySelector(sel); if (!s) return false; return Array.from(s.options).some(o => o.textContent.toLowerCase().replace(/\\s+/g, '').includes(t) || t.includes(o.textContent.toLowerCase().replace(/\\s+/g, ''))); }",
        arg=[sel, t],
        timeout=timeout,
    )


def fuzzy_pick(options, wanted):
    w = norm(wanted)
    for o in options:
        if w in norm(o) or norm(o) in w:
            return o
    return None


def select_fuzzy(page, sel, wanted, label):
    wait_for_option(page, sel, wanted)
    options = options_of(page, sel)
    print(f"{label} options: {options}")
    match = fuzzy_pick(options, wanted)
    if not match:
        save_debug(page, f"no_match_{label.replace(' ', '_').lower()}")
        raise RuntimeError(f"{label}: '{wanted}' not found. Available: {options}")
    page.select_option(sel, label=match)
    page.wait_for_load_state("networkidle")
    return match


def options_of(page, sel):
    return page.eval_on_selector(
        sel,
        "el => Array.from(el.options).map(o => o.textContent.trim()).filter(t => t && !t.toLowerCase().startsWith('select'))"
    )


def click_submit(page):
    """Flexibly find and click the report-generating button."""
    for sel in SUBMIT_SELECTORS:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=1500):
                print(f"Submitting via: {sel}")
                loc.click()
                page.wait_for_load_state("networkidle")
                return
        except Exception:
            continue
    save_debug(page, "no_submit_button")
    raise RuntimeError("No submit button found with any known selector.")


def parse_report(page, metric_name):
    frames = []
    for table in page.query_selector_all("table"):
        html = table.evaluate("el => el.outerHTML")
        try:
            df = pd.read_html(io.StringIO(html))[0]
        except Exception:
            continue
        if df.empty or len(df.columns) < 3:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        date_col = next((c for c in df.columns if "date" in c.lower()), None)
        bo_col = next((c for c in df.columns if c.lower() in ("bo", "bo name", "office_name", "office name") or "branch" in c.lower()), None)
        if not date_col or not bo_col:
            continue
        df = df.rename(columns={date_col: "Date", bo_col: "BO"})
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        df["BO"] = df["BO"].astype(str).str.strip()
        type_col = next((c for c in df.columns if c.lower() in ("type", "office type", "office_type")), None)
        if type_col:
            df = df[df[type_col].astype(str).str.lower().str.contains("bo")]
        num_cols = [c for c in df.columns if c not in ("Date", "BO") and pd.api.types.is_numeric_dtype(df[c])]
        if not num_cols:
            continue
        out = df[["Date", "BO", num_cols[0]]].dropna(subset=["Date"]).copy()
        out.columns = ["Date", "BO", metric_name]
        frames.append(out)
    return frames


def main():
    import io
    all_frames = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            open_page(page)
            page.select_option(CIRCLE_SEL, label=CIRCLE)
            select_fuzzy(page, REGION_SEL, REGION, "Region")
            wait_for_option(page, DIV_SEL, "Division", timeout=60000)
            divisions = options_of(page, DIV_SEL)
            print(f"Divisions: {divisions}")
            for div in divisions:
                page.select_option(DIV_SEL, label=div)
                page.wait_for_function(
                    "sel => { const s = document.querySelector(sel); return s && Array.from(s.options).some(o => !o.textContent.toLowerCase().startsWith('select')); }",
                    arg=SUB_SEL,
                    timeout=60000,
                )
                subs = options_of(page, SUB_SEL)
                for sub in subs:
                    page.select_option(SUB_SEL, label=sub)
                    page.wait_for_load_state("networkidle")
                    for radio_sel, metric in METRICS:
                        page.check(radio_sel)
                        click_submit(page)
                        all_frames.extend(parse_report(page, metric))
                        open_page(page)
                        page.select_option(CIRCLE_SEL, label=CIRCLE)
                        select_fuzzy(page, REGION_SEL, REGION, "Region")
                        wait_for_option(page, DIV_SEL, "Division", timeout=60000)
                        page.select_option(DIV_SEL, label=div)
                        page.wait_for_function(
                            "sel => { const s = document.querySelector(sel); return s && Array.from(s.options).some(o => !o.textContent.toLowerCase().startsWith('select')); }",
                            arg=SUB_SEL,
                            timeout=60000,
                        )
                        page.select_option(SUB_SEL, label=sub)
                        page.wait_for_load_state("networkidle")
        except Exception:
            save_debug(page, "exception")
            traceback.print_exc()
            raise
        finally:
            browser.close()

    if not all_frames:
        sys.exit("No BO-wise data extracted from any report.")
    long_df = pd.concat(all_frames, ignore_index=True)
    metric_cols = [m for _, m in METRICS]
    wide = None
    for mc in metric_cols:
        if mc in long_df.columns:
            part = long_df[["Date", "BO", mc]].dropna(subset=[mc])
            part = part.groupby(["Date", "BO"], as_index=False)[mc].sum()
            wide = part if wide is None else wide.merge(part, on=["Date", "BO"], how="outer")
    for mc in metric_cols:
        if mc not in wide.columns:
            wide[mc] = 0
    wide = wide.fillna(0)
    target = DATA_DIR / "dashboard_data.xlsx"
    wide.to_excel(target, index=False)
    print(f"Combined {len(wide)} BO-wise rows with metrics {metric_cols} -> {target}")


if __name__ == "__main__":
    main()
