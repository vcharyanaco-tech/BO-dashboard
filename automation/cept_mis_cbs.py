"""
CEPT MIS CBSTrans3 BO-wise automation - Excel export flow.

For each division/sub-division and each KPI (Accounts Opened/Closed/Net Addition):
- select radio, submit, wait up to 180s for 'Dated:' signal
- click Excel export button, download to data/exports/{KPI}_{division}_{subdivision}.xlsx

At the end:
- merge all files per KPI into one merged file (3 total)
- delete intermediate per-sub-division files
- upload/use only the 3 merged files for dashboard
"""
import io, os, re, sys, time, traceback, glob
from datetime import date
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

ROOT = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
EXPORT_DIR = DATA_DIR / "exports"
DEBUG_DIR = ROOT / "debug"
DATA_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
DEBUG_DIR.mkdir(exist_ok=True)

REPORT_URL = "https://mis.cept.gov.in/CBS/CBSTrans3.aspx"
CIRCLE = "Haryana Circle"
REGION = "DivReporting Circle"
METRICS = [
    ("#ContentPlaceHolder1_ACO", "accounts_opened"),
    ("#ContentPlaceHolder1_ACC", "accounts_closed"),
    ("#ContentPlaceHolder1_NET", "net_addition"),
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

EXPORT_SELECTORS = [
    "a:has-text('Excel')",
    "button:has-text('Excel')",
    "[value*='Excel']",
    "[onclick*='Excel']",
    "a[title*='Excel']",
    "[id*='Excel']",
    "a:has-text('Export')",
    "button:has-text('Export')",
    "[value*='Export']",
    "[onclick*='Export']",
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


def options_of(page, sel):
    return page.eval_on_selector(
        sel,
        "el => Array.from(el.options).map(o => o.textContent.trim()).filter(t => t && !t.toLowerCase().startsWith('select'))"
    )


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


def wait_for_report_ready(page, timeout=180000):
    page.wait_for_function(
        "() => document.body.innerText.toLowerCase().includes('dated:') || document.body.innerText.toLowerCase().includes('sol id / bocode')",
        timeout=timeout,
    )


def try_click_submit(page):
    for sel in SUBMIT_SELECTORS:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=1500):
                print(f"Submitting via: {sel}")
                loc.click(no_wait_after=True)
                safe_wait(page, timeout=30000)
                return True
        except Exception:
            continue
    save_debug(page, "no_submit_button")
    print("No submit button found.")
    return False


def try_click_export(page, timeout=60000):
    """Click Excel export button, wait for download, return downloaded file path."""
    with page.expect_download(timeout=timeout) as dl_info:
        for sel in EXPORT_SELECTORS:
            try:
                loc = page.locator(sel).first
                if loc.is_visible(timeout=1500):
                    print(f"Exporting via: {sel}")
                    loc.click(no_wait_after=True)
                    break
            except Exception:
                continue
        else:
            save_debug(page, "no_export_button")
            raise RuntimeError("No Excel export button found.")
    download = dl_info.value
    target = EXPORT_DIR / download.suggested_filename
    download.save_as(str(target))
    print(f"Downloaded: {target}")
    return target


def select_radio(page, radio_sel):
    page.click(radio_sel, no_wait_after=True)
    safe_wait(page)


def wait_for_subdivision(page, timeout=90000):
    page.wait_for_function(
        "sel => { const s = document.querySelector(sel); return s && Array.from(s.options).some(o => !o.textContent.toLowerCase().startsWith('select')); }",
        arg=SUB_SEL,
        timeout=timeout,
    )


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


def run_metric_export(page, div, sub, radio_sel, metric, max_attempts=3):
    """Export one KPI report for one sub-division, retry up to max_attempts."""
    for attempt in range(1, max_attempts + 1):
        try:
            print(f"{metric} | {div} / {sub} | attempt {attempt}/{max_attempts}")
            select_radio(page, radio_sel)
            try_click_submit(page)
            print(f"{metric}: waiting up to 180s for report to render...")
            wait_for_report_ready(page, timeout=180000)
            fname = f"{metric}_{norm(div)[:40]}_{norm(sub)[:40]}.xlsx"
            target = EXPORT_DIR / fname
            if target.exists():
                target.unlink()
            try_click_export(page)
            if target.exists() and target.stat().st_size > 0:
                print(f"{metric}: export success -> {target}")
                return target
            print(f"{metric}: export file missing/empty on attempt {attempt}")
        except PWTimeout as e:
            print(f"{metric}: timeout on attempt {attempt}: {e}")
        except Exception as e:
            print(f"{metric}: error on attempt {attempt}: {e}")
        save_debug(page, f"retry_{metric}_{attempt}")
        if attempt < max_attempts:
            time.sleep(5)
            try:
                page.reload(wait_until="domcontentloaded")
                safe_wait(page)
            except Exception:
                pass
            setup_filters(page, div, sub)
    print(f"{metric}: giving up after {max_attempts} attempts for {div}/{sub}")
    return None


def merge_kpi_files(metric):
    """Merge all per-sub-division files for one KPI into one merged file, delete intermediates."""
    pattern = str(EXPORT_DIR / f"{metric}_*.xlsx")
    files = sorted(glob.glob(pattern))
    if not files:
        print(f"{metric}: no files to merge")
        return None
    frames = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df["__source"] = Path(f).stem
            frames.append(df)
        except Exception as e:
            print(f"{metric}: failed to read {f}: {e}")
    if not frames:
        print(f"{metric}: no readable frames")
        return None
    merged = pd.concat(frames, ignore_index=True)
    merged_out = EXPORT_DIR / f"{metric}_merged.xlsx"
    merged.to_excel(merged_out, index=False)
    print(f"{metric}: merged {len(frames)} files -> {merged_out}")
    for f in files:
        try:
            os.remove(f)
            print(f"{metric}: deleted intermediate {f}")
        except Exception as e:
            print(f"{metric}: could not delete {f}: {e}")
    return merged_out


def main():
    merged_files = []
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
                            run_metric_export(page, div, sub, radio_sel, metric)
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

    for _, metric in METRICS:
        out = merge_kpi_files(metric)
        if out:
            merged_files.append(out)

    if not merged_files:
        sys.exit("No merged KPI files produced.")
    print(f"Done. Merged files: {[str(f) for f in merged_files]}")


if __name__ == "__main__":
    main()
