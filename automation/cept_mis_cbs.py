"""
CEPT MIS CBSTrans3 (Product Wise A/C Report) BO-wise automation.

Correct flow:
1. Open https://mis.cept.gov.in/ -> click 'CBS'
2. Same page par Circle/Region/Division/SubDivision dropdowns + 3 radio buttons hain
3. Set End Date = today
4. For each Division -> Sub Division -> radio (Opened/Closed/Net Addition):
   select all, click submit, parse BO-wise rows
5. Merge into wide format: Date, BO, Accounts Opened, Accounts Closed, Net Addition
6. Save data/dashboard_data.xlsx
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

HOME = "https://mis.cept.gov.in/"
CIRCLE = "haryana circle"
REGION = "divreporting circle"
METRICS = [
    ("account opened", "Accounts Opened"),
    ("account closed", "Accounts Closed"),
    ("net addition", "Net Addition"),
]


def save_debug(page, tag):
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{tag}.png"), full_page=True)
        (DEBUG_DIR / f"{tag}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass


def open_report_page(page):
    page.goto(HOME, wait_until="networkidle")
    page.click("text=CBS")
    page.wait_for_load_state("networkidle")


def find_select(page, text):
    for el in page.query_selector_all("select"):
        opts = el.eval_on_selector_all("option", "els => els.map(e => e.textContent.trim().toLowerCase())")
        if any(text in o for o in opts):
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


def select_radio(page, metric_text):
    radios = page.query_selector_all("input[type=radio]")
    for r in radios:
        label_text = r.evaluate(
            "el => (el.labels && el.labels[0] ? el.labels[0].textContent : el.parentElement.textContent || '').trim().toLowerCase()"
        )
        if metric_text in label_text:
            r.check()
            return
    for r in radios:
        val = (r.get_attribute("value") or "").lower()
        if metric_text.split()[0] in val:
            r.check()
            return
    raise RuntimeError(f"Radio not found for: {metric_text}")


def click_submit(page):
    for sel in ["input[type=submit]", "button[type=submit]", "button"]:
        loc = page.locator(sel).first
        try:
            if loc.is_visible(timeout=2000):
                loc.click()
                page.wait_for_load_state("networkidle")
                return
        except Exception:
            continue
    save_debug(page, "no_submit_button")
    raise RuntimeError("No submit button found.")


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
            open_report_page(page)
            end_input = page.query_selector("input[id*='Date'], input[name*='Date'], input[type='text']")
            if end_input:
                end_input.fill(date.today().strftime("%d/%m/%Y"))

            circle_el = find_select(page, CIRCLE)
            region_el = find_select(page, REGION)
            if not circle_el or not region_el:
                save_debug(page, "no_circle_region")
                sys.exit("Circle/Region dropdowns not found on CBS page. Debug artifacts saved.")
            select_by_text(page, circle_el, CIRCLE)
            select_by_text(page, region_el, REGION)

            def get_div_sub():
                div_el = sub_el = None
                for el in page.query_selector_all("select"):
                    if el in (circle_el, region_el):
                        continue
                    opts = [o.lower() for o in options_of(el)]
                    if div_el is None and sum("division" in o for o in opts) >= 3:
                        div_el = el
                    elif div_el is not None and sub_el is None and el != div_el and sum("division" in o or "sub" in o for o in opts) >= 3:
                        sub_el = el
                return div_el, sub_el

            div_el, sub_el = get_div_sub()
            if not div_el or not sub_el:
                save_debug(page, "no_div_sub")
                sys.exit("Division/Sub Division dropdowns not found. Debug artifacts saved.")

            for div in options_of(div_el):
                select_by_text(page, div_el, div)
                for sub in options_of(sub_el):
                    for key, label in METRICS:
                        select_radio(page, key)
                        click_submit(page)
                        all_frames.extend(parse_report(page, label))
                        # reload page and re-select for next metric/sub
                        open_report_page(page)
                        end_input = page.query_selector("input[id*='Date'], input[name*='Date'], input[type='text']")
                        if end_input:
                            end_input.fill(date.today().strftime("%d/%m/%Y"))
                        circle_el = find_select(page, CIRCLE)
                        region_el = find_select(page, REGION)
                        select_by_text(page, circle_el, CIRCLE)
                        select_by_text(page, region_el, REGION)
                        div_el, sub_el = get_div_sub()
        except Exception:
            save_debug(page, "exception")
            traceback.print_exc()
            raise
        finally:
            browser.close()

    if not all_frames:
        sys.exit("No BO-wise data extracted from any report.")
    long_df = pd.concat(all_frames, ignore_index=True)
    metric_cols = [label for _, label in METRICS]
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
