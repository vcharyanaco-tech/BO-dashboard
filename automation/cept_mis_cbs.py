"""
CEPT MIS CBSTrans3 BO-wise automation - Excel export flow.

For each division/sub-division and each KPI (Accounts Opened/Closed/Net Addition):
- select radio, wait up to 90s for 'Dated:' signal (iframe-aware)
- click Excel export button, save download to data/exports/{KPI}_{division}_{subdivision}.xlsx

At the end:
- merge all files per KPI into one merged file (3 total)
- delete intermediate per-sub-division files
- generate data/dashboard_data.xlsx from the 3 merged files for the dashboard
"""
import io, os, re, sys, time, traceback, glob
from datetime import date, timedelta
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
    ("#ContentPlaceHolder1_ACO", "accounts_opened", "Accounts Opened"),
    ("#ContentPlaceHolder1_ACC", "accounts_closed", "Accounts Closed"),
    ("#ContentPlaceHolder1_NET", "net_addition", "Net Addition"),
]

CIRCLE_SEL = "#ContentPlaceHolder1_CircleDropDown"
REGION_SEL = "#ContentPlaceHolder1_RegionDropDown"
DIV_SEL = "#ContentPlaceHolder1_DivisionDropDown"
SUB_SEL = "#ContentPlaceHolder1_SubDivisionDropDown"

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


def log(msg):
    print(msg, flush=True)


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
    page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=60000)
    safe_wait(page)
    yesterday = (date.today() - timedelta(days=1)).strftime("%d/%m/%Y")
    date_inputs = page.query_selector_all("input[id*='Date'], input[name*='Date'], input[type='text']")
    for di in date_inputs[:2]:
        try:
            di.fill(yesterday)
        except Exception:
            pass


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
    page.select_option(sel, label=match, timeout=60000)
    safe_wait(page)
    return match


def report_ready(page):
    for frame in page.frames:
        try:
            txt = frame.evaluate("() => document.body ? document.body.innerText.toLowerCase() : ''")
            if 'dated:' in txt or 'sol id / bocode' in txt:
                return True
        except Exception:
            continue
    return False


def wait_for_report_ready(page, timeout=90000, refresh_retries=2):
    deadline = time.time() + timeout
    refreshes = 0
    while time.time() < deadline:
        if report_ready(page):
            return True
        time.sleep(5)
        if time.time() >= deadline:
            break
        if refreshes < refresh_retries and (deadline - time.time()) < timeout * 0.5:
            log("Report not ready - refreshing page and re-checking...")
            refreshes += 1
            try:
                page.reload(wait_until="domcontentloaded", timeout=60000)
                safe_wait(page)
            except Exception as e:
                log(f"Reload failed: {e}")
    return report_ready(page)


def try_click_export(page, target, timeout=60000):
    """Click Excel export and save download to the intended target path."""
    with page.expect_download(timeout=timeout) as dl_info:
        for sel in EXPORT_SELECTORS:
            try:
                loc = page.locator(sel).first
                if loc.is_visible(timeout=1500):
                    log(f"Exporting via: {sel}")
                    loc.click(no_wait_after=True, timeout=10000)
                    break
            except Exception:
                continue
        else:
            save_debug(page, "no_export_button")
            raise RuntimeError("No Excel export button found.")
    download = dl_info.value
    download.save_as(str(target))
    log(f"Downloaded: {target} (suggested name: {download.suggested_filename})")
    return target


def select_radio(page, radio_sel):
    page.click(radio_sel, no_wait_after=True, timeout=10000)
    safe_wait(page)


def wait_for_subdivision(page, timeout=60000):
    page.wait_for_function(
        "sel => { const s = document.querySelector(sel); return s && Array.from(s.options).some(o => !o.textContent.toLowerCase().startsWith('select')); }",
        arg=SUB_SEL,
        timeout=timeout,
    )


def setup_filters(page, div=None, sub=None):
    open_page(page)
    page.select_option(CIRCLE_SEL, label=CIRCLE, timeout=60000)
    select_fuzzy(page, REGION_SEL, REGION, "Region")
    wait_for_option(page, DIV_SEL, "Division", timeout=60000)
    if div:
        page.select_option(DIV_SEL, label=div, timeout=60000)
    if sub:
        wait_for_subdivision(page, timeout=60000)
        page.select_option(SUB_SEL, label=sub, timeout=60000)
    safe_wait(page)


def run_metric_export(page, div, sub, radio_sel, metric_key, metric_label, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            log(f"{metric_key} | {div} / {sub} | attempt {attempt}/{max_attempts}")
            select_radio(page, radio_sel)
            log(f"{metric_key}: waiting up to 90s for report to render (iframe-aware, with refresh retry)...")
            if wait_for_report_ready(page, timeout=90000):
                fname = f"{metric_key}_{norm(div)[:40]}_{norm(sub)[:40]}.xlsx"
                target = EXPORT_DIR / fname
                if target.exists():
                    target.unlink()
                try_click_export(page, target)
                if target.exists() and target.stat().st_size > 0:
                    log(f"{metric_key}: export success -> {target}")
                    return target
                log(f"{metric_key}: export file missing/empty on attempt {attempt}")
            else:
                log(f"{metric_key}: report never rendered on attempt {attempt}")
        except PWTimeout as e:
            log(f"{metric_key}: timeout on attempt {attempt}: {e}")
        except Exception as e:
            log(f"{metric_key}: error on attempt {attempt}: {e}")
        save_debug(page, f"retry_{metric_key}_{attempt}")
        if attempt < max_attempts:
            time.sleep(5)
            try:
                page.reload(wait_until="domcontentloaded", timeout=60000)
                safe_wait(page)
            except Exception:
                pass
            setup_filters(page, div, sub)
    log(f"{metric_key}: giving up after {max_attempts} attempts for {div}/{sub}")
    return None


def merge_kpi_files(metric_key):
    pattern = str(EXPORT_DIR / f"{metric_key}_*.xlsx")
    files = [f for f in sorted(glob.glob(pattern)) if "merged" not in f]
    if not files:
        log(f"{metric_key}: no files to merge")
        return None
    frames = []
    for f in files:
        try:
            df = pd.read_excel(f, header=None)
            frames.append(df)
        except Exception as e:
            log(f"{metric_key}: failed to read {f}: {e}")
    if not frames:
        log(f"{metric_key}: no readable frames")
        return None
    merged = pd.concat(frames, ignore_index=True)
    merged_out = EXPORT_DIR / f"{metric_key}_merged.xlsx"
    merged.to_excel(merged_out, index=False)
    log(f"{metric_key}: merged {len(frames)} files -> {merged_out}")
    for f in files:
        try:
            os.remove(f)
            log(f"{metric_key}: deleted intermediate {f}")
        except Exception as e:
            log(f"{metric_key}: could not delete {f}: {e}")
    return merged_out


def generate_dashboard_data():
    today = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    per_metric = {}
    for _, metric_key, metric_label in METRICS:
        merged_file = EXPORT_DIR / f"{metric_key}_merged.xlsx"
        if not merged_file.exists():
            log(f"{metric_key}: merged file missing, skipping")
            continue
        try:
            matrix = pd.read_excel(merged_file, header=None).values.tolist()
        except Exception as e:
            log(f"{metric_key}: failed to read merged file: {e}")
            continue
        # header detection
        hIdx = -1
        for i, row in enumerate(matrix):
            cells = [str(c).strip() if c is not None else "" for c in row]
            if any(c.lower() in ("name", "bo", "bo name", "branch office") or "sol id" in c.lower() or "bocode" in c.lower() for c in cells):
                hIdx = i
                break
        if hIdx < 0:
            log(f"{metric_key}: header row not found in merged file")
            continue
        raw_headers = [str(c).strip() if c is not None else "" for c in matrix[hIdx]]
        seen = {}
        headers = []
        for c in raw_headers:
            c = c or "col"
            if c in seen:
                seen[c] += 1
                headers.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                headers.append(c)
        rows = [r for r in matrix[hIdx+1:] if any(str(c).strip() for c in r if c is not None)]
        df = pd.DataFrame(rows, columns=headers)
        bo_pos = next((i for i, h in enumerate(headers) if h.lower() in ("name", "bo", "bo name", "branch office") or "sol id" in h.lower() or "bocode" in h.lower()), None)
        if bo_pos is None:
            log(f"{metric_key}: no BO column found")
            continue
        bo_series = df.iloc[:, bo_pos].astype(str).str.strip()
        num_pos = []
        for i in range(len(headers)):
            if i == bo_pos:
                continue
            coerced = pd.to_numeric(df.iloc[:, i], errors="coerce")
            if coerced.notna().any():
                num_pos.append(i)
        if not num_pos:
            log(f"{metric_key}: no numeric columns found")
            continue
        metric_vals = df.iloc[:, num_pos].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
        part = pd.DataFrame({"BO": bo_series, metric_label: metric_vals}).groupby("BO", as_index=False)[metric_label].sum()
        per_metric[metric_label] = part
    if not per_metric:
        log("No metric data to build dashboard file.")
        return None
    combined = None
    for label, part in per_metric.items():
        combined = part if combined is None else combined.merge(part, on="BO", how="outer")
    combined.insert(0, "Date", today)
    for _, _, metric_label in METRICS:
        if metric_label not in combined.columns:
            combined[metric_label] = 0
    combined = combined.fillna(0)
    out = DATA_DIR / "dashboard_data.xlsx"
    combined.to_excel(out, index=False)
    log(f"Dashboard data written: {out} ({len(combined)} BO rows)")
    return out


def main():
    merged_files = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            setup_filters(page)
            divisions = options_of(page, DIV_SEL)
            log(f"Divisions: {divisions}")
            for div in divisions:
                try:
                    page.select_option(DIV_SEL, label=div, timeout=60000)
                    wait_for_subdivision(page, timeout=60000)
                except PWTimeout:
                    log(f"{div}: subdivision dropdown timeout - reloading and retrying once")
                    save_debug(page, f"sub_timeout_{norm(div)[:30]}")
                    try:
                        setup_filters(page, div)
                        wait_for_subdivision(page, timeout=60000)
                    except PWTimeout:
                        log(f"{div}: still timing out - skipping this division")
                        continue
                subs = options_of(page, SUB_SEL)
                for sub in subs:
                    try:
                        page.select_option(SUB_SEL, label=sub, timeout=60000)
                        safe_wait(page)
                        for radio_sel, metric_key, metric_label in METRICS:
                            run_metric_export(page, div, sub, radio_sel, metric_key, metric_label)
                    except PWTimeout as e:
                        log(f"{div}/{sub}: timeout ({e}) - reloading and moving to next sub")
                        save_debug(page, f"sub_select_timeout_{norm(sub)[:30]}")
                        try:
                            setup_filters(page, div)
                        except Exception:
                            pass
                        continue
                    except Exception as e:
                        log(f"{div}/{sub}: error ({e}) - reloading and moving to next sub")
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

    for _, metric_key, _ in METRICS:
        out = merge_kpi_files(metric_key)
        if out:
            merged_files.append(out)

    dashboard_file = generate_dashboard_data()
    if not dashboard_file:
        sys.exit("No dashboard data file produced.")
    log(f"Done. Merged files: {[str(f) for f in merged_files]}")
    log(f"Dashboard data: {dashboard_file}")


if __name__ == "__main__":
    main()
