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
- The Accounts Opened radio `#ContentPlaceHolder1_ACO` (name Group1) fires a
  postback that re-renders the result grid.
- Export is client-side: `dom: 'Bifrtp'` + `{extend: 'excel', footer: true}` means
  the button is `.dt-buttons button.buttons-excel` inside `#example1_wrapper`.
  There is no server-side export URL; the workbook is built in the browser and
  its filename comes from document.title ("Product Wise A/C Report"), which is
  why the download is saved under our own target name.

Per division/sub-division and KPI the export is saved to
data/exports/{KPI}_{division}_{subdivision}.xlsx and validated to contain real BO
rows. At the end the per-sub-division files are merged per KPI into
data/exports/{kpi}_merged.xlsx (one header row + data rows only),
generate_dashboard_data() builds data/dashboard_data.xlsx for the dashboard,
append_history() appends the run to data/history.csv (daily snapshots that feed
the dashboard's trend chart), and write_manifest() writes
data/exports/manifest.json (the KPI list the dashboard auto-loads).
"""
import os
import re
import sys
import time
import glob
import json
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

ROOT = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
EXPORT_DIR = DATA_DIR / "exports"
HISTORY_FILE = DATA_DIR / "history.csv"
DEBUG_DIR = ROOT / "debug"
DATA_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
DEBUG_DIR.mkdir(exist_ok=True)

REPORT_URL = "https://mis.cept.gov.in/CBS/CBSTrans3.aspx"
CIRCLE = os.getenv("REPORT_CIRCLE", "Haryana Circle")
REGION = os.getenv("REPORT_REGION", "DivReportingCircle")
# "subdivision" (default) or "division": division runs return the same BO rows
# with ~4x fewer page loads, which the fragile server tolerates much better.
REPORT_LEVEL = os.getenv("REPORT_LEVEL", "subdivision").strip().lower()

# Page date inputs are DD-MM-YYYY (datepicker "dd-mm-yy");
# the "Data Updated upto" label is DD/MM/YYYY.
DATE_FMT = "%d-%m-%Y"
LABEL_DATE_FMT = "%d/%m/%Y"

# Only Accounts Opened is fetched: the dashboard reports that single KPI.
METRICS = [
    ("#ContentPlaceHolder1_ACO", "accounts_opened", "Accounts Opened"),
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


def cell_text(value):
    """Cell -> stripped text. Excel blanks arrive as NaN, so treat them as empty."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


def safe_wait(page, timeout=6000):
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


SERVER_ERROR_RE = re.compile(r"server error in|timeout expired|pooled connections", re.I)


def page_is_error(page):
    """True when MIS answered with its ASP.NET 'Server Error' page (a dropped DB
    connection under load) instead of the report - the run has to reload."""
    try:
        text = str(page.title()) + " " + str(page.eval_on_selector("body", "el => el.innerText || ''"))
    except Exception:
        return True
    return bool(SERVER_ERROR_RE.search(text))


def open_page(page, attempts=4):
    """Load the report page, retrying the site's intermittent server errors.

    Without this the whole run dies on the first dropped connection: the filter
    dropdowns simply never appear on the error page, so select_option() waits out
    its 60s timeout and the fetch for the day is lost.
    """
    for i in range(1, attempts + 1):
        page.goto(REPORT_URL, wait_until="domcontentloaded", timeout=90000)
        safe_wait(page)
        if not page_is_error(page):
            return
        log(f"report page served a server error (attempt {i}/{attempts}) - reloading")
        save_debug(page, f"server_error_{i}")
        if i < attempts:
            time.sleep(20 * i)
    raise RuntimeError(f"CEPT MIS kept returning its 'Server Error' page after {attempts} loads")


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
  // The CEPT page reuses several DataTables and moves the rendered report between
  // them: Accounts Opened/Closed land in #example1 while Net Addition lands in
  // #example4. Pick the visible table that actually holds rows instead of assuming.
  const $ = window.jQuery;
  const cands = Array.from(document.querySelectorAll('table.dataTable, table[id^="example"]'));
  let best = {id: null, rows: -1};
  for (const t of cands) {
    if (!t.offsetParent) continue;                       // hidden table
    let rows = 0;
    try {
      if ($ && $.fn.dataTable && $.fn.dataTable.isDataTable(t)) {
        rows = $(t).DataTable().rows().count();
      }
    } catch (e) { rows = 0; }
    if (!rows) {
      for (const tr of t.querySelectorAll('tr')) {
        const tds = tr.querySelectorAll('td');
        if (tds.length < 3) continue;                    // header/footer cells are <th>
        if ((tds[0].innerText || '').trim().toLowerCase().startsWith('total')) continue;
        rows++;
      }
    }
    if (rows > best.rows) best = {id: t.id || null, rows: rows};
  }
  return best;
}
"""


def current_grid(page):
    """Visible report table as {'id': ..., 'rows': ...}."""
    try:
        grid = page.evaluate(GRID_ROW_JS)
    except Exception:
        grid = None
    if not isinstance(grid, dict):
        return {"id": None, "rows": -1}
    return {"id": grid.get("id"), "rows": int(grid.get("rows", -1))}


def wait_for_grid(page, timeout=90000, min_rows=1):
    """Poll until the rendered report has real BO rows (fresh page has no grid at all)."""
    deadline = time.time() + timeout
    grid = {"id": None, "rows": -1}
    while time.time() < deadline:
        grid = current_grid(page)
        if grid["rows"] >= min_rows:
            return grid
        time.sleep(2)
    return grid


def select_radio(page, radio_sel):
    """Tick a KPI radio and fire its postback without blocking on the navigation.

    page.click() waits for the navigation its click triggers, and the CEPT server can
    leave that request hanging forever (observed on Net Addition), which froze the whole
    sweep. Dispatching from JS keeps every wait bounded.
    """
    page.evaluate(
        """sel => {
            const el = document.querySelector(sel);
            if (!el) throw new Error('radio not found: ' + sel);
            el.click();
        }""",
        radio_sel,
    )
    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
    except PWTimeout:
        pass
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
            if is_header_row([cell_text(c) for c in row]):
                h_idx = i
                break
        if h_idx < 0:
            return False, 0
        valid = 0
        for row in matrix[h_idx + 1:]:
            cells = [cell_text(c) for c in row]
            if not any(cells):
                continue
            if is_header_row(cells) or is_footer_row(cells) or is_title_row(cells):
                continue
            valid += 1
        return valid > 0, valid
    except Exception as e:
        log(f"has_bo_data_rows error: {e}")
        return False, 0


def export_selectors(grid_id=None):
    """Excel button for the rendered grid first, generic ones only as a last resort."""
    scoped = []
    if grid_id:
        scoped = [
            f"#{grid_id}_wrapper .dt-buttons button.buttons-excel",
            f"#{grid_id}_wrapper .dt-buttons .buttons-excel",
            f"#{grid_id}_wrapper button.buttons-excel",
        ]
    fallback = list(EXPORT_SELECTORS) if not grid_id else [
        "#example1_wrapper .dt-buttons button.buttons-excel",
        "button:has-text('Excel')",
    ]
    return scoped + [sel for sel in fallback if sel not in scoped]


def try_click_export(page, target, grid_id=None, timeout=90000):
    with page.expect_download(timeout=timeout) as dl_info:
        used = None
        for sel in export_selectors(grid_id):
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
    target = EXPORT_DIR / f"{metric_key}_{norm(div)[:40]}_{norm(sub)[:40]}.xlsx"
    if os.getenv("SKIP_EXISTING", "1") == "1" and target.exists():
        has_data, row_count = has_bo_data_rows(target)
        if has_data:
            log(f"{metric_key}: already downloaded -> {target.name} ({row_count} BO rows)")
            return target
    for attempt in range(1, max_attempts + 1):
        try:
            log(f"{metric_key} | {div} / {sub} | attempt {attempt}/{max_attempts} | window {start_s} -> {end_s}")
            select_radio(page, radio_sel)
            grid = wait_for_grid(page, timeout=90000)
            rows = grid["rows"]

            dated = dated_window(page)
            if dated and dated != (start_s, end_s):
                log(f"{metric_key}: report caption says {dated[0]} - {dated[1]}, applying {start_s} - {end_s}")
                ensure_window(page, window)
                click_submit(page)
                select_radio(page, radio_sel)
                grid = wait_for_grid(page, timeout=90000)
                rows = grid["rows"]
                dated = dated_window(page)
            if dated:
                log(f"{metric_key}: report caption {dated[0]} - {dated[1]}")

            if rows < 1:
                if dated == (start_s, end_s):
                    # Report rendered for the right window with zero BOs - a valid empty
                    # result (e.g. no net change). Retrying would just burn 3x90s.
                    log(f"{metric_key}: 0 BO rows in {start_s} -> {end_s}; nothing to export")
                    return None
                log(f"{metric_key}: grid rendered no BO rows on attempt {attempt}")
            else:
                log(f"{metric_key}: {rows} BO rows rendered")
                if target.exists():
                    target.unlink()
                try_click_export(page, target, grid.get("id"))
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
        cells = [cell_text(c) for c in row]
        if is_header_row(cells):
            keep.append(not header_kept)
            header_kept = True
            continue
        if not any(cells) or is_footer_row(cells) or is_title_row(cells):
            keep.append(False)
            continue
        keep.append(True)

    cleaned = merged.loc[[i for i, keep_row in enumerate(keep) if keep_row]].reset_index(drop=True)
    merged_out = EXPORT_DIR / f"{metric_key}_merged.xlsx"
    cleaned.to_excel(merged_out, index=False, header=False)
    log(f"{metric_key}: merged {len(frames)} files -> {merged_out} ({len(cleaned) - 1} BO rows)")
    for f in files:
        try:
            os.remove(f)
            log(f"{metric_key}: deleted intermediate {f}")
        except Exception as e:
            log(f"{metric_key}: could not delete {f}: {e}")
    return merged_out


SO_COL_RE = re.compile(r"^so\s*name$|^sub\s*office|^so$", re.I)
SOL_COL_RE = re.compile(r"sol\s*id|bocode|bo\s*code", re.I)
# Same office-name and office-code patterns as app.js, so the snapshot and the
# dashboard classify the interleaved office summary rows identically.
SO_NAME_RE = re.compile(
    r"\b(s\.?\s?o\.?|h\.?\s?o\.?|p\.?\s?o\.?|g\.?\s?p\.?\s?o\.?|head\s*office|sub\s*office)\b", re.I)
OFFICE_CODE_RE = re.compile(r"^\d{6,}$")  # SO/HO codes are 8 digits; BO codes are 5 alphanumeric
BO_SUFFIX_RE = re.compile(r"\s*B\.?O\.?$", re.I)


def norm_bo_name(name):
    """Report/master BO name in its canonical comparison form (mirrors app.js normName)."""
    s = " ".join(str(name or "").split())
    return BO_SUFFIX_RE.sub(" BO", s).strip().lower()


def load_bo_master():
    """Office identity from bo_master.js: codes (SOL_ID_BO_ID) and names.

    The code is the key - the master sheet has 2,196 offices but only 2,094 distinct
    names, so 82 names belong to two offices each and a name alone cannot identify one.
    """
    try:
        txt = (ROOT / "bo_master.js").read_text(encoding="utf-8")
        m = re.search(r"BO_MASTER\s*=\s*(\[.*\])\s*;?\s*$", txt, re.S)
        if not m:
            return set(), set()
        codes, names = set(), set()
        for o in json.loads(m.group(1)):
            code = str(o.get("code") or "").strip().upper()
            if code:
                codes.add(code)
            name = norm_bo_name(o.get("name"))
            if name:
                names.add(name)
        return codes, names
    except Exception as e:
        log(f"bo_master: could not load master list ({e}) - S.O filter falls back to name pattern only")
        return set(), set()


BO_MASTER_CODES, BO_MASTER_NAMES = load_bo_master()


def is_master_office(bo, sol=""):
    """True when the row is a BO listed in the master sheet - code first, then name."""
    code = str(sol or "").strip().upper()
    if code and code in BO_MASTER_CODES:
        return True
    return norm_bo_name(bo) in BO_MASTER_NAMES


def is_so_row(bo, sol=""):
    """True for the office summary rows interleaved in the CEPT grid ('Adhoya S.O',
    'Jamalpur SO Bhiwani', 'Ambala G.P.O.'): never a master BO, and either an
    office-style name or an all-digit office code."""
    if is_master_office(bo, sol):
        return False
    return bool(SO_NAME_RE.search(str(bo or "").strip())) or bool(OFFICE_CODE_RE.match(str(sol or "").strip().upper()))


def so_key(name):
    """Identity space for the office summary rows (mirrors app.js soKey: dots go)."""
    return "so:" + " ".join(str(name or "").replace(".", "").split()).lower()


def row_identity(kind, bo, sol=""):
    """One entry per OFFICE, not per name: SOL ID / BOCODE first, name only as a
    fallback. Two different BOs can share a name (the master sheet has 82 such names),
    so keying by name would silently add their figures together and drop one row."""
    if kind == "so":
        return so_key(bo)
    return str(sol or "").strip().upper() or norm_bo_name(bo)


def _read_merged(metric_key):
    """Read one merged KPI workbook -> {office id: {bo, sol, so, value, kind}}."""
    merged_file = EXPORT_DIR / f"{metric_key}_merged.xlsx"
    if not merged_file.exists():
        log(f"{metric_key}: merged file missing, skipping")
        return {}
    try:
        matrix = pd.read_excel(merged_file, header=None).values.tolist()
    except Exception as e:
        log(f"{metric_key}: failed to read merged file: {e}")
        return {}
    h_idx = -1
    for i, row in enumerate(matrix):
        if is_header_row([cell_text(c) for c in row]):
            h_idx = i
            break
    if h_idx < 0:
        log(f"{metric_key}: header row not found in merged file")
        return {}

    headers = [cell_text(c) or f"col{i}" for i, c in enumerate(matrix[h_idx])]
    bo_pos = next((i for i, h in enumerate(headers)
                   if h.lower() in HEADER_HINTS or "sol id" in h.lower() or "bocode" in h.lower()), None)
    if bo_pos is None:
        log(f"{metric_key}: no BO column found")
        return {}
    sol_pos = next((i for i, h in enumerate(headers) if SOL_COL_RE.search(h)), None)
    so_pos = next((i for i, h in enumerate(headers) if SO_COL_RE.search(h)), None)

    body = matrix[h_idx + 1:]
    num_pos = []
    for i in range(len(headers)):
        if i in (bo_pos, sol_pos):
            continue
        coerced = pd.to_numeric(
            pd.Series([cell_text(r[i]) if i < len(r) else "" for r in body]).str.replace(",", "", regex=False),
            errors="coerce",
        )
        if coerced.notna().any():
            num_pos.append(i)
    if not num_pos:
        log(f"{metric_key}: no numeric columns found")
        return {}

    out = {}
    for r in body:
        cells = [cell_text(c) for c in r]
        if not any(cells) or is_footer_row(cells) or is_title_row(cells) or is_header_row(cells):
            continue
        bo = cells[bo_pos] if bo_pos < len(cells) else ""
        if not bo or bo.lower().startswith("total"):
            continue
        value = 0.0
        for i in num_pos:
            if i < len(cells):
                try:
                    value += float(cells[i].replace(",", "") or 0)
                except ValueError:
                    pass
        so_cell = cells[so_pos] if (so_pos is not None and so_pos < len(cells)) else ""
        sol_cell = cells[sol_pos] if (sol_pos is not None and sol_pos < len(cells)) else ""
        # Office summary rows (S.O/H.O/G.P.O: blank SO Name cell or an office-style
        # name/code, and not a known master BO) carry SO-LEVEL figures - keep them
        # separate as kind='so' so the dashboard never mixes them into BO sums.
        office = is_so_row(bo, sol_cell) or (so_cell == "" and not is_master_office(bo, sol_cell))
        kind = "so" if office else "bo"
        key = row_identity(kind, bo, sol_cell)
        entry = out.setdefault(key, {"bo": bo, "sol": "", "so": "", "value": 0.0, "kind": kind})
        entry["value"] += value
        if not entry["sol"] and sol_cell:
            entry["sol"] = sol_cell
        if kind == "so":
            entry["so"] = bo  # the office row groups under its own name
        elif not entry["so"] and so_cell:
            entry["so"] = so_cell
    return out


def generate_dashboard_data(report_date_iso=None):
    """Build data/dashboard_data.xlsx: one row per BO with SO, SOL ID and KPI values.

    report_date_iso is the end of the reporting window (the report's own date),
    falling back to yesterday when unknown.
    """
    report_date = report_date_iso or (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    metric_values = {}
    meta = {}
    for _, metric_key, metric_label in METRICS:
        data = _read_merged(metric_key)
        # dashboard_data.xlsx stays BO-only; SO-level rows (kind='so') are excluded
        # here - the dashboard reads those from history.csv/exports, kept separate.
        data = {k: v for k, v in data.items() if v.get("kind", "bo") != "so"}
        if not data:
            continue
        metric_values[metric_label] = {k: v["value"] for k, v in data.items()}
        for k, v in data.items():
            m = meta.setdefault(k, {"bo": v["bo"], "sol": "", "so": ""})
            m["sol"] = m["sol"] or v["sol"]
            m["so"] = m["so"] or v["so"]

    if not metric_values:
        log("No metric data to build dashboard file.")
        return None

    labels = [m[2] for m in METRICS]
    ids = sorted(meta)
    frame = pd.DataFrame({
        "Date": report_date,
        "BO": [meta[i]["bo"] for i in ids],
        "SO": [meta[i]["so"] for i in ids],
        "SOL ID": [meta[i]["sol"] for i in ids],
    })
    for label in labels:
        values = metric_values.get(label)
        if values is None:
            log(f"{label}: no data in this run, filled with 0")
            values = {}
        frame[label] = [values.get(i, 0) for i in ids]
    out = DATA_DIR / "dashboard_data.xlsx"
    frame.to_excel(out, index=False)
    with_so = sum(1 for s in frame["SO"] if s)
    log(f"Dashboard data written: {out} ({len(frame)} BO rows, {with_so} with SO names)")
    return out


def append_history(report_date_iso):
    """Append this run's BO-wise KPI values to data/history.csv (long format).

    Columns: date,bo,so,sol,kpi,value — one row per BO per KPI per report date.
    Idempotent per report date: re-running for the same date replaces that date's
    rows instead of duplicating them, so the dashboard's cumulative trend chart
    never double-counts a re-run.
    """
    if not report_date_iso:
        log("history: no report date given - skipped")
        return HISTORY_FILE if HISTORY_FILE.exists() else None
    rows = []
    for _, metric_key, metric_label in METRICS:
        data = _read_merged(metric_key)
        for _id, v in data.items():
            rows.append({
                "date": report_date_iso,
                "bo": v["bo"],
                "so": v.get("so", ""),
                "sol": v.get("sol", ""),
                "kind": v.get("kind", "bo"),
                "kpi": metric_label,
                "value": v.get("value", 0),
            })
    if not rows:
        log("history: no metric data this run - history file left unchanged")
        return HISTORY_FILE if HISTORY_FILE.exists() else None
    new = pd.DataFrame(rows)
    if HISTORY_FILE.exists():
        try:
            old = pd.read_csv(HISTORY_FILE, dtype=str)
            for col in ("date", "bo", "so", "sol", "kind", "kpi", "value"):
                if col not in old.columns:
                    old[col] = "" if col != "kind" else "bo"
            old = old[old["date"] != report_date_iso]
            combined = pd.concat([old, new], ignore_index=True)
        except Exception as e:
            log(f"history: could not read existing {HISTORY_FILE} ({e}) - rewriting it")
            combined = new
    else:
        combined = new
    for col in ("date", "bo", "so", "sol", "kind", "kpi"):
        combined[col] = combined[col].fillna("" if col != "kind" else "bo").astype(str)
    # One row per office per day, identified by code (name only when there is none).
    # Deduplicating by name would merge the 82 master names that two BOs share.
    combined["_id"] = [row_identity(k, b, s) for k, b, s in
                       zip(combined["kind"], combined["bo"], combined["sol"])]
    combined = combined.drop_duplicates(subset=["date", "_id", "kind", "kpi"], keep="last")
    combined = combined.drop(columns=["_id"])
    combined = combined.sort_values(["date", "kpi", "bo", "sol"]).reset_index(drop=True)
    vals = pd.to_numeric(combined["value"], errors="coerce").fillna(0)
    if (vals % 1 == 0).all():
        vals = vals.astype("int64")
    combined["value"] = vals
    combined.to_csv(HISTORY_FILE, index=False)
    log(f"history: appended {len(new)} rows for {report_date_iso} -> {HISTORY_FILE} "
        f"(total {len(combined)} rows, {combined['date'].nunique()} day(s))")
    return HISTORY_FILE


def write_manifest(report_date_iso=None):
    """Write data/exports/manifest.json listing the merged KPI workbooks to auto-load.

    The dashboard fetches this manifest instead of a hardcoded KPI list, so adding
    a KPI to METRICS needs no frontend change. Each entry carries the report date
    because the merged workbooks no longer keep their 'Dated:' title row.
    """
    manifest = []
    for _, metric_key, metric_label in METRICS:
        f = EXPORT_DIR / f"{metric_key}_merged.xlsx"
        if not f.exists():
            log(f"manifest: {f.name} missing - excluded")
            continue
        has_data, _ = has_bo_data_rows(f)
        if not has_data:
            log(f"manifest: {f.name} has no BO data rows - excluded")
            continue
        entry = {"file": f"data/exports/{f.name}", "label": metric_label}
        if report_date_iso:
            entry["date"] = report_date_iso
        manifest.append(entry)
    out = EXPORT_DIR / "manifest.json"
    out.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    log(f"manifest: wrote {out} ({len(manifest)} KPI source(s))")
    return out


def main():
    merged_files = []
    report_iso = None
    launch_args = ["--no-sandbox", "--disable-dev-shm-usage"] if os.getenv("CI") else []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=launch_args)
        page = browser.new_page(accept_downloads=True)
        page.set_default_timeout(25000)
        page.set_default_navigation_timeout(45000)
        try:
            open_page(page)
            updated, label = read_updated_label(page)
            start, end = resolve_window(updated)
            report_iso = end.isoformat()
            window = (start.strftime(DATE_FMT), end.strftime(DATE_FMT))
            log(f"Data window: {window[0]} -> {window[1]}")
            if label:
                log(f"Page label: {label}")
            if updated:
                log(f"Data updated upto: {updated.strftime(DATE_FMT)}")

            setup_filters(page, window)
            divisions = options_of(page, DIV_SEL)
            only = [d.strip().lower() for d in os.getenv("REPORT_DIVISIONS", "").split(",") if d.strip()]
            if only:
                divisions = [d for d in divisions if any(o in d.lower() for o in only)]
                log(f"REPORT_DIVISIONS filter -> {divisions}")
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
                if REPORT_LEVEL == "division":
                    subs = ["All"]
                    log(f"{div}: division-level run (one export per KPI)")
                else:
                    subs = options_of(page, SUB_SEL)
                    log(f"{div}: {len(subs)} sub divisions")
                for sub in subs:
                    try:
                        if sub != "All":
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

    if os.getenv("SKIP_MERGE", "") == "1":
        log("SKIP_MERGE=1: merge + dashboard build left to the caller")
        return

    for _, metric_key, _ in METRICS:
        out = merge_kpi_files(metric_key)
        if out:
            merged_files.append(out)

    if not merged_files:
        # Nothing was downloaded this run. The merged files on disk are the previous
        # run's, and history/manifest are built from them - so without this guard a
        # failed scrape silently files yesterday's (or last week's) numbers under
        # today's date and the dashboard then adds them up as extra days.
        sys.exit("No KPI workbook was downloaded for this window. Refusing to rewrite "
                 "history/dashboard/manifest from a stale merged file - check debug/ and re-run.")

    dashboard_file = generate_dashboard_data(report_iso)
    history_file = append_history(report_iso)
    manifest_file = write_manifest(report_iso)
    if not dashboard_file:
        sys.exit("No dashboard data file produced.")
    log(f"Done. Merged files: {[str(f) for f in merged_files]}")
    log(f"Dashboard data: {dashboard_file}")
    log(f"History: {history_file}")
    log(f"Manifest: {manifest_file}")


if __name__ == "__main__":
    main()
