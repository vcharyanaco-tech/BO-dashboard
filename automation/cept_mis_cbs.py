"""
CEPT MIS CBSTrans3 BO-wise automation.

Report structure (per radio selection):
Name | SOL ID / BOCODE | SO Name | MIS | PPFGP | SSA | RD | SBBAS | SBSGP | SCSS | TD | PRFTS | KVN | NSC8 | MSSC

BO = Name column; metric = sum of all numeric product columns; Total row skipped; Date = today.
"""
import io, sys, time, traceback
from datetime import date
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

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
    "input[value*='Product Wise']",
    "a:has-text('Product Wise')",
    "[value*='Report']",
    "[onclick*='Product']",
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


def safe_wait(page, timeout=20000):
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except PWTimeout:
        pass


def save_debug(page, tag):
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{tag}.png"), full_page=True)
        (DEBUG_DIR / f"{tag}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass


def open_page(page):
    page.goto(REPORT_URL, wait_until="domcontentloaded")
    safe_wait(page)
    date_input = page.query_selector("input[id*='Date'], input[name*='Date'], input[type='text']")
    if date_input:
        date_input.fill(date.today().strftime("%d/%m/%Y"))


def wait_for_option(page, sel, text, timeout=90000):
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
    match = fuzzy_pick(options, wanted)
    if not match:
        save_debug(page, f"no_match_{label.replace(' ', '_').lower()}")
        raise RuntimeError(f"{label}: '{wanted}' not found. Available: {options}")
    page.select_option(sel, label=match)
    safe_wait(page)
    return match


def options_of(page, sel):
    return page.eval_on_selector(
        sel,
        "el => Array.from(el.options).map(o => o.textContent.trim()).filter(t => t && !t.toLowerCase().startsWith('select'))"
    )


def try_click_submit(page):
    for sel in SUBMIT_SELECTORS:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=1500):
                print(f"Submitting via: {sel}")
                loc.click(no_wait_after=True)
                safe_wait(page, timeout=45000)
                return True
        except Exception:
            continue
    save_debug(page, "no_submit_button")
    print("No submit button found.")
    return False


def select_radio(page, radio_sel):
    page.click(radio_sel, no_wait_after=True)
    safe_wait(page)


def log_tables(page, context):
    tables = page.query_selector_all("table")
    print(f"[{context}] Found {len(tables)} tables")
    for i, table in enumerate(tables):
        try:
            html = table.evaluate("el => el.outerHTML")
            df = pd.read_html(io.StringIO(html))[0]
            print(f"[{context}] Table {i}: {len(df)} rows, columns: {list(df.columns)[:10]}")
        except Exception as e:
            print(f"[{context}] Table {i}: unreadable ({e})")


def parse_report(page, metric_name):
    """Parse Product Wise A/C Report: Name = BO, sum numeric product columns = metric."""
    today = date.today().strftime("%Y-%m-%d")
    frames = []
    for table in page.query_selector_all("table"):
        html = table.evaluate("el => el.outerHTML")
        try:
            df = pd.read_html(io.StringIO(html))[0]
        except Exception:
            continue
        if df.empty or len(df.columns) < 4:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        name_col = next((c for c in df.columns if norm(c) in ("name", "officename", "bo") or "name" in norm(c)), None)
        sol_col = next((c for c in df.columns if "solar" in norm(c) or "solarid" in norm(c) or "bocode" in norm(c)), None)
        if not name_col or not sol_col:
            continue
        df = df.rename(columns={name_col: "BO"})
        df["BO"] = df["BO"].astype(str).str.strip()
        # Skip Total row
        df = df[~df["BO"].str.lower().str.startswith("total")]
        df = df[df["BO"] != ""]
        num_cols = [c for c in df.columns if c not in ("BO", name_col, sol_col) and pd.api.types.is_numeric_dtype(df[c])]
        if not num_cols:
            continue
        df[metric_name] = df[num_cols].sum(axis=1)
        out = df[["BO", metric_name]].copy()
        out.insert(0, "Date", today)
        frames.append(out)
    return frames


def setup_filters(page, div=None, sub=None):
    open_page(page)
    page.select_option(CIRCLE_SEL, label=CIRCLE)
    select_fuzzy(page, REGION_SEL, REGION, "Region")
    wait_for_option(page, DIV_SEL, "Division", timeout=90000)
    if div:
        page.select_option(DIV_SEL, label=div)
    if sub:
        wait_for_subdivision(page, timeout=90000)
        page.select_option(SUB_SEL, label=sub)
    safe_wait(page)


def wait_for_subdivision(page, timeout=90000):
    page.wait_for_function(
        "sel => { const s = document.querySelector(sel); return s && Array.from(s.options).some(o => !o.textContent.toLowerCase().startsWith('select')); }",
        arg=SUB_SEL,
        timeout=timeout,
    )


def run_metric_with_retry(page, div, sub, radio_sel, metric, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            print(f"{metric} | {div} / {sub} | attempt {attempt}/{max_attempts}")
            select_radio(page, radio_sel)
            try_click_submit(page)
            frames = parse_report(page, metric)
            if frames:
                return frames
            print(f"{metric}: empty parse on attempt {attempt}, logging tables...")
            log_tables(page, f"{metric} attempt {attempt}")
        except PWTimeout as e:
            print(f"{metric}: timeout on attempt {attempt}: {e}")
        except Exception as e:
            print(f"{metric}: error on attempt {attempt}: {e}")
        save_debug(page, f"retry_{metric.replace(' ', '_').lower()}_{attempt}")
        if attempt < max_attempts:
            time.sleep(5)
            try:
                page.reload(wait_until="domcontentloaded")
                safe_wait(page)
            except Exception:
                pass
            setup_filters(page, div, sub)
    print(f"{metric}: giving up after {max_attempts} attempts for {div}/{sub}")
    return []


def main():
    import io
    all_frames = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            setup_filters(page)
            divisions = options_of(page, DIV_SEL)
            print(f"Divisions: {divisions}")
            for div in divisions:
                try:
                    page.select_option(DIV_SEL, label=div)
                    wait_for_subdivision(page, timeout=90000)
                except PWTimeout:
                    print(f"{div}: subdivision dropdown timeout - reloading and retrying once")
                    save_debug(page, f"sub_timeout_{norm(div)[:30]}")
                    try:
                        setup_filters(page, div)
                        wait_for_subdivision(page, timeout=90000)
                    except PWTimeout:
                        print(f"{div}: still timing out - skipping this division")
                        continue
                subs = options_of(page, SUB_SEL)
                for sub in subs:
                    try:
                        page.select_option(SUB_SEL, label=sub)
                        safe_wait(page)
                        for radio_sel, metric in METRICS:
                            all_frames.extend(run_metric_with_retry(page, div, sub, radio_sel, metric))
                    except PWTimeout as e:
                        print(f"{div}/{sub}: timeout ({e}) - reloading and moving to next sub")
                        save_debug(page, f"sub_select_timeout_{norm(sub)[:30]}")
                        try:
                            setup_filters(page, div)
                        except Exception:
                            pass
                        continue
                    except Exception as e:
                        print(f"{div}/{sub}: error ({e}) - reloading and moving to next sub")
                        save_debug(page, f"sub_select_error_{norm(sub)[:30]}")
                        try:
                            setup_filters(page, div)
                        except Exception:
                            pass
                        continue
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
