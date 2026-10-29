"""
CEPT MIS CBSTrans3 (Product Wise A/C Report) BO-wise automation.

Cascading dropdowns load via ASP.NET postbacks, so after each selection we wait
until the next dropdown actually has options before selecting.
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


def wait_options(page, sel, min_options=2, timeout=30000):
    page.wait_for_function(
        "([sel, n]) => document.querySelector(sel) && document.querySelector(sel).options.length >= n",
        arg=[sel, min_options],
        timeout=timeout,
    )


def select_circle_region(page):
    page.select_option(CIRCLE_SEL, label=CIRCLE)
    wait_options(page, REGION_SEL)          # Region options load after Circle postback
    page.select_option(REGION_SEL, label=REGION)
    wait_options(page, DIV_SEL)             # Division options load after Region postback


def options_of(page, sel):
    return page.eval_on_selector(
        sel,
        "el => Array.from(el.options).map(o => o.textContent.trim()).filter(t => t && !t.toLowerCase().startsWith('select'))"
    )


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
            select_circle_region(page)
            divisions = options_of(page, DIV_SEL)
            for div in divisions:
                page.select_option(DIV_SEL, label=div)
                wait_options(page, SUB_SEL)     # SubDivision options load after Division postback
                subs = options_of(page, SUB_SEL)
                for sub in subs:
                    page.select_option(SUB_SEL, label=sub)
                    page.wait_for_load_state("networkidle")
                    for radio_sel, metric in METRICS:
                        page.check(radio_sel)
                        page.click("input[type=submit]")
                        page.wait_for_load_state("networkidle")
                        all_frames.extend(parse_report(page, metric))
                        # reload fresh form for next metric/sub
                        open_page(page)
                        select_circle_region(page)
                        page.select_option(DIV_SEL, label=div)
                        wait_options(page, SUB_SEL)
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
