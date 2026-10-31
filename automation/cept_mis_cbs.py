"""
CEPT MIS CBSTrans3 BO-wise automation - Excel export flow.

Workflow as verified against the live page (https://mis.cept.gov.in/CBS/CBSTrans3.aspx):

- Classic ASP.NET WebForms page (__VIEWSTATE / __EVENTVALIDATION). Every filter
  change is a full __doPostBack, so each step has to wait for the next set of
  real options before selecting.
- Dates are two text inputs (`startDate`, `endDate`) backed by a jquery-ui
  datepicker with dateFormat "dd-mm-yy" => DD-MM-YYYY. The page header prints
  "Data Updated upto: DD/MM/YYYY"; asking for a window beyond that date renders
  a header-only report, so the window is always clamped to it.
- Cascade: Circle -> Region -> Division -> Sub Division.
- The "Submit" button (`button.btn-danger`, __doPostBack('ctl00$ContentPlaceHolder1$ctl00',''))
  renders the report. The grid is table `#example1`:
      Name | SOL ID / BOCODE | SO Name | MIS PPFGP SSA RD SBBAS SBSGP SCSS TD PRFTS KVN NSC8 MSSC
  Data rows are plain `<tr>` elements (there is NO <tbody>), and `<tfoot>` holds
  a "Total:" row that DataTables fills in and the Excel export includes
  (button config uses `footer: true`).
- KPI radios `#ContentPlaceHolder1_ACO` / `_ACC` / `_NET` (name Group1) each fire
  a postback that re-renders `#example1` for that KPI.
- Export is client-side: `dom: 'Bifrtp'` + `{extend: 'excel', footer: true}` means
  the button is `.dt-buttons button.buttons-excel` inside `#example1_wrapper`.
  There is no server-side export URL; the workbook is built in the browser and
  its filename comes from document.title ("Product Wise A/C Report"), which is
  why the download is saved under our own target name.

Per division/sub-division and KPI the export is saved to
data/exports/{KPI}_{division}_{subdivision}.xlsx and validated to contain real BO
rows. At the end the per-sub-division files are merged per KPI into
data/exports/{kpi}_merged.xlsx (one header row + data rows only) and
generate_dashboard_data() builds data/dashboard_data.xlsx for the dashboard.
"""
import os
import re
import sys
import time
import glob
import traceback
from datetime import date, datetime, timedelta
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
CIRCLE = os.getenv("REPORT_CIRCLE", "Haryana Circle")
REGION = os.getenv("REPORT_REGION", "DivReportingCircle")

# Page date inputs are DD-MM-YYYY (datepicker "dd-mm-yy");
# the "Data Updated upto" label is DD/MM/YYYY.
DATE_FMT = "%d-%m-%Y"
LABEL_DATE_FMT = "%d/%m/%Y"

METRICS = [
    ("#ContentPlaceHolder1_ACO", "accounts_opened", "Accounts Opened"),
    ("#ContentPlaceHolder1_ACC", "accounts_closed", "Accounts Closed"),
    ("#ContentPlaceHolder1_NET", "net_addition", "Net Addition"),
]

CIRCLE_SEL = "#ContentPlaceHolder1_CircleDropDown"
REGION_SEL = "#ContentPlaceHolder1_RegionDropDown"
DIV_SEL = "#ContentPlaceHolder1_DivisionDropDown"
SUB_SEL = "#ContentPlaceHolder1_SubDivisionDropDown"
START_SEL = "#ContentPlaceHolder1_startDate"
END_SEL = "#ContentPlaceHolder1_endDate"
UPDATED_LABEL_SEL = "#ContentPlaceHolder1_Label2"
GRID_SEL = "#example1"

# The grid's own DataTables Excel button lives inside the table wrapper; keep the
# scoped selectors first so a different table's export button is never clicked.
EXPORT_SELECTORS = [
    "#example1_wrapper .dt-buttons button.buttons-excel",
    "#example1_wrapper .dt-buttons .buttons-excel",
    ".dt-buttons button.buttons-excel",
    "button.buttons-excel",
    "button:has-text('Excel')",
]

# Header row of the BO grid / of the exported workbook.
HEADER_HINTS = ("name", "bo", "bo name", "branch office")


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


# --------------------------------------------------------------------------- #
# Dates
# --------------------------------------------------------------------------- #
def read_updated_label(page):
    """Return the raw "Data Updated upto: ..." label text (already rendered server-side)."""
    txt = ""
    try:
        txt = page.eval_on_selector(UPDATED_LABEL_SEL, "el => el.textContent || ''")
    except Exception:
        txt = ""
    if not txt.strip():
        try:
            txt = page.evaluate("() => document.body ? document.body.innerText : ''")
        except Exception:
            txt = ""
    m = re.search(r"Data Updated upto:?\s*(\d{2})/(\d{2})/(\d{4})", txt, re.I)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))), m.group(0).strip()
        except ValueError:
            pass
    return None, txt.strip()[:100]


def resolve_window(updated_upto):
    """Pick the reporting window.

    Defaults to a single previous day (what the dashboard has always used), but
    never later than the date the page says its data is updated to, because a
    future window renders an empty report. REPORT_START_DATE / REPORT_END_DATE
    (DD-MM-YYYY, or DD/MM/YYYY) override either end.
    """
    def parse(value, env_name):
        for fmt in (DATE_FMT, LABEL_DATE_FMT):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
        raise SystemExit(f"{env_name}='{value}' is not a valid date; use DD-MM-YYYY")

    end_env = os.getenv("REPORT_END_DATE", "").strip()
    start_env = os.getenv("REPORT_START_DATE", "").strip()

    end = parse(end_env, "REPORT_END_DATE") if end_env else (date.today() - timedelta(days=1))
    if updated_upto and end > updated_upto:
        end = updated_upto

    start = parse(start_env, "REPORT_START_DATE") if start_env else end
    if start > end:
        start = end
    return start, end


def set_date(page, sel, value, label):
    """Set a datepicker-backed input and confirm the value stuck."""
    try:
        page.eval_on_selector(
            sel,
            "(el, v) => { el.value = v;"
            " el.dispatchEvent(new Event('change', { bubbles: true }));"
            " el.dispatchEvent(new Event('blur', { bubbles: true })); }",
            value,
        )
    except Exception:
        page.fill(sel, value)
    try:
        page.keyboard.press("Escape")  # dismiss the datepicker overlay
    except Exception:
        pass
    actual = page.eval_on_selector(sel, "el => (el.value || '').trim()")
    if str(actual) != value:
        save_debug(page, f"date_mismatch_{norm(label)[:20]}")
        raise RuntimeError(f"{label}: field kept {actual!r} instead of {value!r}")


def ensure_window(page, window):
    """Make sure both date inputs carry the window; re-set them after postbacks."""
    start_s, end_s = window
    for sel, val, name in ((START_SEL, start_s, "startDate"), (END_SEL, end_s, "endDate")):
        try:
            current = page.eval_on_selector(sel, "el => (el.value || '').trim()")
        except Exception:
            current = None
        if current != val:
            set_date(page, sel, val, name)
    return window


def dated_window(page):
    """Read the report's own "Dated: DD-MM-YYYY - DD-MM-YYYY" caption."""
    try:
        txt = page.evaluate("() => document.body ? document.body.innerText : ''")
    except Exception:
        return None
    m = re.search(r"Dated:\s*(\d{2}-\d{2}-\d{4})\s*-\s*(\d{2}-\d{2}-\d{4})", txt)
    return (m.group(1), m.group(2)) if m else None


# --------------------------------------------------------------------------- #
# Filters
# --------------------------------------------------------------------------- #
def options_of(page, sel):
    return page.eval_on_selector(
        sel,
        "el => Array.from(el.options).map(o => o.textContent.trim()).filter(t => t && !t.toLowerCase().startsWith('select'))",
    )


def wait_for_options(page, sel, timeout=60000):
    """Wait until a dropdown has at least one real option (not a 'Select ...' placeholder)."""
    page.wait_for_function(
        "sel => { const s = document.querySelector(sel); if (!s) return false;"
        " return Array.from(s.options).some(o => { const t = (o.textContent || '').trim().toLowerCase();"
        " return t && !t.startsWith('select'); }); }",
        arg=sel,
        timeout=timeout,
    )


def wait_for_subdivision(page, timeout=60000):
    wait_for_options(page, SUB_SEL, timeout=timeout)


def wait_for_option(page, sel, text, timeout=60000):
    t = norm(text)
    page.wait_for_function(
        "([sel, t]) => { const s = document.querySelector(sel); if (!s) return false;"
        " return Array.from(s.options).some(o => { const v = o.textContent.toLowerCase().replace(/\\s+/g, '');"
        " return v.includes(t) || t.includes(v); }); }",
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
    page.select_option(sel, label=match, timeout=60000)
    safe_wait(page)
    return match


def open_page(page):
    page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=90000)
    safe_wait(page)


def setup_filters(page, window, div=None, sub=None):
    open_page(page)
    page.select_option(CIRCLE_SEL, label=CIRCLE, timeout=60000)
    select_fuzzy(page, REGION_SEL, REGION, "Region")
    wait_for_options(page, DIV_SEL, timeout=60000)
    ensure_window(page, window)
    if div:
        page.select_option(DIV_SEL, label=div, timeout=60000)
    if sub:
        wait_for_subdivision(page, timeout=60000)
        page.select_option(SUB_SEL, label=sub, timeout=60000)
    safe_wait(page)
    ensure_window(page, window)


def click_submit(page):
    """Press the page's Submit button (renders the report for the current filters/dates)."""
    clicked = False
    for sel in ("button:has-text('Submit')", "input[value='Submit']", "button.btn-danger"):
        try:
            loc = page.locator(sel)
            if loc.count() and loc.first.is_visible(timeout=1500):
                log(f"Submit via: {sel}")
                loc.first.click(no_wait_after=True, timeout=15000)
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        log("Submit via: __doPostBack fallback")
        page.evaluate("__doPostBack('ctl00$ContentPlaceHolder1$ctl00','')")
    safe_wait(page)
    return clicked


# --------------------------------------------------------------------------- #
# Grid + export
# --------------------------------------------------------------------------- #
GRID_ROW_JS = """
() => {
  const t = document.querySelector('#example1');
  if (!t) return -1;
  // DataTables paginates the grid, so prefer its own row count over the DOM rows.
  try {
    const $ = window.jQuery;
    if ($ && $.fn && $.fn.dataTable && $.fn.dataTable.isDataTable('#example1')) {
      return $('#example1').DataTable().rows().count();
    }
  } catch (e) { /* fall through to DOM counting */ }
  let n = 0;
  for (const tr of t.querySelectorAll('tr')) {
    const tds = tr.querySelectorAll('td');
    if (tds.length < 3) continue;                      // header/footer cells are <th>
    const first = (tds[0].innerText || '').trim().toLowerCase();
    if (first.startsWith('total')) continue;           // DataTables footer row
    n++;
  }
  return n;
}
"""


def grid_row_count(page):
    try:
        n = page.evaluate(GRID_ROW_JS)
        return -1 if n is None else int(n)
    except Exception:
        return -1


def wait_for_grid(page, timeout=90000, min_rows=1):
    """Poll until #example1 has real BO rows (fresh page has no grid at all)."""
    deadline = time.time() + timeout
    rows = -1
    while time.time() < deadline:
        rows = grid_row_count(page)
        if rows >= min_rows:
            return rows
        time.sleep(2)
    return rows


def select_radio(page, radio_sel):
    page.click(radio_sel, no_wait_after=True, timeout=15000)
    safe_wait(page)


def is_footer_row(cells):
    return bool(cells) and cells[0].strip().lower().startswith("total")


def is_title_row(cells):
    filled = [c for c in cells if c]
    if len(filled) != 1:
        return False
    low = filled[0].strip().lower()
    return bool(re.search(r"report|dated:", low))


def is_header_row(cells):
    for c in cells:
        low = c.lower()
        if low in HEADER_HINTS or "sol id" in low or "bocode" in low:
            return True
    return False


def has_bo_data_rows(xlsx_path):
    """True if the workbook contains real BO data rows (not just header/Total)."""
    try:
        df = pd.read_excel(xlsx_path, header=None)
        matrix = df.values.tolist()
        h_idx = -1
        for i, row in enumerate(matrix):
            if is_header_row([str(c).strip() if c is not None else "" for c in row]):
                h_idx = i
                break
        if h_idx < 0:
            return False, 0
        valid = 0
        for row in matrix[h_idx + 1:]:
            cells = [str(c).strip() if c is not None else "" for c in row]
            if not any(cells):
                continue
            if is_header_row(cells) or is_footer_row(cells) or is_title_row(cells):
                continue
            valid += 1
        return valid > 0, valid
    except Exception as e:
        log(f"has_bo_data_rows error: {e}")
        return False, 0


def try_click_export(page, target, timeout=90000):
    with page.expect_download(timeout=timeout) as dl_info:
        used = None
        for sel in EXPORT_SELECTORS:
            try:
                loc = page.locator(sel)
                if loc.count() and loc.first.is_visible(timeout=1500):
                    log(f"Exporting via: {sel}")
                    loc.first.click(no_wait_after=True, timeout=15000)
                    used = sel
                    break
            except Exception:
                continue
        if used is None:
            save_debug(page, "no_export_button")
            raise RuntimeError("No Excel export button found (#example1 .dt-buttons button.buttons-excel).")
    download = dl_info.value
    download.save_as(str(target))
    log(f"Downloaded: {target} (suggested name: {download.suggested_filename})")
    return target


def run_metric_export(page, div, sub, radio_sel, metric_key, metric_label, window, max_attempts=3):
    start_s, end_s = window
    for attempt in range(1, max_attempts + 1):
        try:
            log(f"{metric_key} | {div} / {sub} | attempt {attempt}/{max_attempts} | window {start_s} -> {end_s}")
            select_radio(page, radio_sel)
            rows = wait_for_grid(page, timeout=90000)

            dated = dated_window(page)
            if dated and dated != (start_s, end_s):
                log(f"{metric_key}: report caption says {dated[0]} - {dated[1]}, applying {start_s} - {end_s}")
                ensure_window(page, window)
                click_submit(page)
                select_radio(page, radio_sel)
                rows = wait_for_grid(page, timeout=90000)
                dated = dated_window(page)
            if dated:
                log(f"{metric_key}: report caption {dated[0]} - {dated[1]}")

            if rows < 1:
                log(f"{metric_key}: grid rendered no BO rows on attempt {attempt}")
            else:
                log(f"{metric_key}: {rows} BO rows rendered")
                target = EXPORT_DIR / f"{metric_key}_{norm(div)[:40]}_{norm(sub)[:40]}.xlsx"
                if target.exists():
                    target.unlink()
                try_click_export(page, target)
                if target.exists() and target.stat().st_size > 0:
                    has_data, row_count = has_bo_data_rows(target)
                    if has_data:
                        log(f"{metric_key}: export success -> {target} ({row_count} BO rows)")
                        return target
                    log(f"{metric_key}: export has no BO data rows (header only) on attempt {attempt} - will retry")
                else:
                    log(f"{metric_key}: export file missing/empty on attempt {attempt}")
        except PWTimeout as e:
            log(f"{metric_key}: timeout on attempt {attempt}: {e}")
        except Exception as e:
            log(f"{metric_key}: error on attempt {attempt}: {e}")
        save_debug(page, f"retry_{metric_key}_{attempt}")
        if attempt < max_attempts:
            time.sleep(5)
            try:
                setup_filters(page, window, div, sub)
                click_submit(page)
            except Exception:
                pass
    log(f"{metric_key}: giving up after {max_attempts} attempts for {div}/{sub} - no valid BO data exported")
    return None


# --------------------------------------------------------------------------- #
# Merge + dashboard data
# --------------------------------------------------------------------------- #
def merge_kpi_files(metric_key):
    """Merge the per-sub-division exports of one KPI into a single clean workbook.

    Repeated title/header rows and the DataTables "Total:" footer row are dropped,
    so the merged file holds exactly one header row followed by BO data rows.
    """
    pattern = str(EXPORT_DIR / f"{metric_key}_*.xlsx")
    files = [f for f in sorted(glob.glob(pattern)) if "merged" not in f]
    if not files:
        log(f"{metric_key}: no files to merge")
        return None
    frames = []
    for f in files:
        try:
            has_data, _ = has_bo_data_rows(f)
            if not has_data:
                log(f"{metric_key}: skipping {f} - no BO data rows")
                continue
            frames.append(pd.read_excel(f, header=None))
        except Exception as e:
            log(f"{metric_key}: failed to read {f}: {e}")
    if not frames:
        log(f"{metric_key}: no readable frames with data")
        return None

    merged = pd.concat(frames, ignore_index=True)
    keep = []
    header_kept = False
    for _, row in merged.iterrows():
        cells = [str(c).strip() if c is not None else "" for c in row]
        if is_header_row(cells):
            keep.append(not header_kept)
            header_kept = True
            continue
        if not any(cells) or is_footer_row(cells) or is_title_row(cells):
            keep.append(False)
            continue
        keep.append(True)

    cleaned = merged[keep].reset_index(drop=True)
    merged_out = EXPORT_DIR / f"{metric_key}_merged.xlsx"
    cleaned.to_excel(merged_out, index=False)
    log(f"{metric_key}: merged {len(frames)} files -> {merged_out} ({len(cleaned) - 1} BO rows)")
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
        h_idx = -1
        for i, row in enumerate(matrix):
            if is_header_row([str(c).strip() if c is not None else "" for c in row]):
                h_idx = i
                break
        if h_idx < 0:
            log(f"{metric_key}: header row not found in merged file")
            continue

        raw_headers = [str(c).strip() if c is not None else "" for c in matrix[h_idx]]
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

        rows = [
            r for r in matrix[h_idx + 1:]
            if any(str(c).strip() for c in r if c is not None)
            and not is_footer_row([str(c).strip() if c is not None else "" for c in r])
            and not is_title_row([str(c).strip() if c is not None else "" for c in r])
        ]
        df = pd.DataFrame(rows, columns=headers)
        bo_pos = next(
            (i for i, h in enumerate(headers)
             if h.lower() in HEADER_HINTS or "sol id" in h.lower() or "bocode" in h.lower()),
            None,
        )
        if bo_pos is None:
            log(f"{metric_key}: no BO column found")
            continue

        bo_series = df.iloc[:, bo_pos].astype(str).str.strip()
        # Skip repeated header rows, the "Total:" footer and blank BO cells.
        low = bo_series.str.lower()
        mask = ~low.isin(["name", "bo", "bo name"]) & ~low.str.startswith("total") & (bo_series != "")
        sol_pos = next((i for i, h in enumerate(headers) if "sol id" in h.lower() or "bocode" in h.lower()), None)

        num_pos = []
        for i in range(len(headers)):
            if i == bo_pos or i == sol_pos:
                continue
            coerced = pd.to_numeric(
                df.iloc[:, i].astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce"
            )
            if coerced.notna().any():
                num_pos.append(i)
        if not num_pos:
            log(f"{metric_key}: no numeric columns found")
            continue

        metric_vals = (
            df.iloc[:, num_pos]
            .apply(lambda s: pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce"))
            .fillna(0)
            .sum(axis=1)
        )
        part = (
            pd.DataFrame({"BO": bo_series[mask], metric_label: metric_vals[mask]})
            .groupby("BO", as_index=False)[metric_label]
            .sum()
        )
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
    launch_args = ["--no-sandbox", "--disable-dev-shm-usage"] if os.getenv("CI") else []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=launch_args)
        page = browser.new_page(accept_downloads=True)
        page.set_default_timeout(60000)
        try:
            open_page(page)
            updated, label = read_updated_label(page)
            start, end = resolve_window(updated)
            window = (start.strftime(DATE_FMT), end.strftime(DATE_FMT))
            log(f"Data window: {window[0]} -> {window[1]}")
            if label:
                log(f"Page label: {label}")
            if updated:
                log(f"Data updated upto: {updated.strftime(DATE_FMT)}")

            setup_filters(page, window)
            divisions = options_of(page, DIV_SEL)
            log(f"Divisions ({len(divisions)}): {divisions}")
            for div in divisions:
                try:
                    page.select_option(DIV_SEL, label=div, timeout=60000)
                    wait_for_subdivision(page, timeout=60000)
                except PWTimeout:
                    log(f"{div}: subdivision dropdown timeout - reloading and retrying once")
                    save_debug(page, f"sub_timeout_{norm(div)[:30]}")
                    try:
                        setup_filters(page, window, div)
                        wait_for_subdivision(page, timeout=60000)
                    except PWTimeout:
                        log(f"{div}: still timing out - skipping this division")
                        continue
                subs = options_of(page, SUB_SEL)
                log(f"{div}: {len(subs)} sub divisions")
                for sub in subs:
                    try:
                        page.select_option(SUB_SEL, label=sub, timeout=60000)
                        safe_wait(page)
                        ensure_window(page, window)
                        click_submit(page)
                        for radio_sel, metric_key, metric_label in METRICS:
                            run_metric_export(page, div, sub, radio_sel, metric_key, metric_label, window)
                    except PWTimeout as e:
                        log(f"{div}/{sub}: timeout ({e}) - reloading and moving to next sub")
                        save_debug(page, f"sub_select_timeout_{norm(sub)[:30]}")
                        try:
                            setup_filters(page, window, div)
                        except Exception:
                            pass
                        continue
                    except Exception as e:
                        log(f"{div}/{sub}: error ({e}) - reloading and moving to next sub")
                        save_debug(page, f"sub_select_error_{norm(sub)[:30]}")
                        try:
                            setup_filters(page, window, div)
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
