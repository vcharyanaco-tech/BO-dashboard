"""
CEPT MIS CBSTrans3 (Product Wise A/C Report) BO-wise automation.

Flow:
1. Open https://mis.cept.gov.in/CBS/CBSTrans3.aspx
2. Set End Date = today
3. Select Circle = Haryana Circle, Region = DivReporting Circle
4. For each Division -> each Sub Division: click 'Product Wise A/C Report'
5. Parse rendered HTML table, keep BO-wise rows
6. Combine all into data/dashboard_data.xlsx
"""
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

MIS_URL = "https://mis.cept.gov.in/CBS/CBSTrans3.aspx"
CIRCLE = "haryana circle"
REGION = "divreporting circle"


def find_select(page, option_text):
    """Find a select whose options contain the given text (case-insensitive)."""
    handle = page.query_selector_all("select")
    for el in handle:
        opts = el.eval_on_selector_all("option", "els => els.map(e => e.textContent.trim().toLowerCase())")
        if any(option_text in o for o in opts):
            return el
    return None


def select_by_text(page, el, text):
    value = el.eval_on_selector_all(
        "option",
        "(els,t) => { const m = els.find(e => e.textContent.trim().toLowerCase().includes(t)); return m ? m.value : null; }",
        text.lower(),
    )
    if value is None:
        raise RuntimeError(f"Option not found: {text}")
    el.select_option(value)
    page.wait_for_load_state("networkidle")


def options_of(el):
    return el.eval_on_selector_all(
        "option",
        "els => els.map(e => e.textContent.trim()).filter(t => t && !t.toLowerCase().startsWith('select'))"
    )


def parse_report(page):
    """Parse the rendered report table(s) and return BO-wise rows."""
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
        # Keep only BO rows (exclude SO-wise rows if a type column exists)
        type_col = next((c for c in df.columns if c.lower() in ("type", "office type", "office_type")), None)
        if type_col:
            df = df[df[type_col].astype(str).str.lower().str.contains("bo")]
        keep = ["Date", "BO"] + [c for c in df.columns if c not in ("Date", "BO") and pd.api.types.is_numeric_dtype(df[c])]
        if len(keep) > 2:
            frames.append(df[keep].dropna(subset=["Date"]))
    return frames


def main():
    import io
    all_frames = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(MIS_URL, wait_until="networkidle")

        # End Date = today
        end_input = page.query_selector("input[type='text'][id*='Date'], input[id*='EndDate'], input[name*='Date']")
        if end_input:
            end_input.fill(date.today().strftime("%d/%m/%Y"))

        circle_el = find_select(page, CIRCLE)
        region_el = find_select(page, REGION)
        if not circle_el or not region_el:
            sys.exit("Circle/Region dropdown not found on page.")
        select_by_text(page, circle_el, CIRCLE)
        select_by_text(page, region_el, REGION)

        # Division and Sub Division dropdowns appear after Region postback
        div_el = None
        sub_el = None
        for el in page.query_selector_all("select"):
            opts = [o.lower() for o in options_of(el)]
            if div_el is None and sum("division" in o for o in opts) >= 3 and el != circle_el and el != region_el:
                div_el = el
            elif div_el is not None and sub_el is None and el != div_el and sum("division" in o or "sub" in o for o in opts) >= 3:
                sub_el = el
        if not div_el or not sub_el:
            sys.exit("Division/Sub Division dropdowns not found.")

        for div in options_of(div_el):
            select_by_text(page, div_el, div)
            for sub in options_of(sub_el):
                select_by_text(page, sub_el, sub)
                page.click("input[type=submit][value*='Product Wise A/C Report']")
                page.wait_for_load_state("networkidle")
                all_frames.extend(parse_report(page))
                # go back to report form for next selection
                page.goto(MIS_URL, wait_until="networkidle")
                circle_el = find_select(page, CIRCLE)
                region_el = find_select(page, REGION)
                select_by_text(page, circle_el, CIRCLE)
                select_by_text(page, region_el, REGION)
                div_el = next(el for el in page.query_selector_all("select") if el != circle_el and el != region_el and sum("division" in o.lower() for o in options_of(el)) >= 3)
                sub_el = next(el for el in page.query_selector_all("select") if el not in (circle_el, region_el, div_el) and sum("division" in o.lower() or "sub" in o.lower() for o in options_of(el)) >= 3)
        browser.close()

    if not all_frames:
        sys.exit("No BO-wise data extracted from any report.")
    out = pd.concat(all_frames, ignore_index=True)
    out = out.groupby(["Date", "BO"], as_index=False).sum(numeric_only=True)
    target = DATA_DIR / "dashboard_data.xlsx"
    out.to_excel(target, index=False)
    print(f"Combined {len(out)} BO-wise rows -> {target}")


if __name__ == "__main__":
    main()
