"""
CEPT MIS CBSTrans3 BO-wise report automation.

- Opens https://mis.cept.gov.in/CBS/CBSTrans3.aspx
- Selects Circle=Haryana Circle, Region=DivReporting Circle (fixed)
- Iterates every Division and each of its Sub Divisions
- Downloads each report, extracts BO-wise rows
- Combines everything into data/dashboard_data.xlsx for the dashboard

Run locally:  python automation/cept_mis_cbs.py
GitHub Actions runs this daily and commits the result.
"""
import io, os, sys
from datetime import date
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

MIS_URL = "https://mis.cept.gov.in/CBS/CBSTrans3.aspx"
CIRCLE = "Haryana Circle"
REGION = "DivReporting Circle"


def select_by_text(page, selector, text):
    """Select dropdown option whose visible text matches (case-insensitive)."""
    page.select_option(selector, label=text)


def get_options(page, selector):
    return page.eval_on_selector(
        selector,
        "el => Array.from(el.options).map(o => o.text.trim()).filter(t => t && t.toLowerCase() !== 'select')"
    )


def find_selects(page):
    """Heuristically identify Circle/Region/Division/SubDivision dropdowns by their options."""
    selects = page.eval_on_selector_all(
        "select",
        "els => els.map((el,i) => ({i, id: el.id, name: el.name, options: Array.from(el.options).map(o => o.text.trim())}))"
    )
    circle = next((s for s in selects if any(CIRCLE.lower() in o.lower() for o in s["options"])), None)
    region = next((s for s in selects if any(REGION.lower() in o.lower() for o in s["options"]) and s["i"] != circle["i"]), None)
    # Division dropdown: contains 'Division' in many option texts
    division = next((s for s in selects if s["i"] not in (circle["i"], region["i"]) and sum('division' in o.lower() for o in s["options"]) >= 3), None)
    subdiv = next((s for s in selects if s["i"] not in (circle["i"], region["i"], division["i"]) and sum('sub' in o.lower() or 'division' in o.lower() for o in s["options"]) >= 3), None)
    return circle, region, division, subdiv


def sel(page, meta):
    return f"select#{meta['id']}" if meta["id"] else f"select[name='{meta['name']}']"


def main():
    all_frames = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(accept_downloads=True)
        page.goto(MIS_URL, wait_until="networkidle")

        circle, region, division, subdiv = find_selects(page)
        if not all([circle, region, division, subdiv]):
            sys.exit("Could not auto-detect Circle/Region/Division/SubDivision dropdowns. Page structure may have changed.")

        select_by_text(page, sel(page, circle), CIRCLE)
        page.wait_for_load_state("networkidle")
        select_by_text(page, sel(page, region), REGION)
        page.wait_for_load_state("networkidle")

        divisions = get_options(page, sel(page, division))
        for div in divisions:
            select_by_text(page, sel(page, division), div)
            page.wait_for_load_state("networkidle")
            subs = get_options(page, sel(page, subdiv))
            for sub in subs:
                select_by_text(page, sel(page, subdiv), sub)
                page.wait_for_load_state("networkidle")
                # Trigger report generation/download
                try:
                    with page.expect_download(timeout=30000) as dl:
                        page.click("input[type=submit][value*='Report'], a[href*='download'], input[value*='Show'], input[value*='Generate']")
                    path = dl.value.path()
                    df = pd.read_excel(path) if not path.endswith(".csv") else pd.read_csv(path)
                except Exception:
                    # Fallback: report may render as HTML table instead of file download
                    try:
                        df = pd.read_html(page.content())[0]
                    except Exception:
                        continue
                df.columns = [str(c).strip() for c in df.columns]
                date_col = next((c for c in df.columns if "date" in c.lower()), None)
                bo_col = next((c for c in df.columns if c.lower() in ("bo", "bo name", "office_name", "office name") or "branch" in c.lower()), None)
                if not date_col or not bo_col:
                    continue
                df = df.rename(columns={date_col: "Date", bo_col: "BO"})
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
                df["BO"] = df["BO"].astype(str).str.strip()
                keep = ["Date", "BO"] + [c for c in df.columns if c not in ("Date", "BO") and pd.api.types.is_numeric_dtype(df[c])]
                all_frames.append(df[keep].dropna(subset=["Date"]))
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
