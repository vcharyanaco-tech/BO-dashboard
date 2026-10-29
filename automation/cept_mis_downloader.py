"""
CEPT MIS BO-wise report downloader & consolidator.

- Logs into CEPT MIS using credentials from automation/.env
- Downloads BO-wise reports (configurable in automation/config.json)
- Consolidates all reports into a single data/dashboard_data.xlsx
- Optionally pushes the consolidated file to GitHub so the dashboard auto-loads it

Usage:
    python cept_mis_downloader.py
"""
import json, os, subprocess, sys
from datetime import date
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

BASE = Path(__file__).parent
ROOT = BASE.parent
DATA_DIR = ROOT / "data"
DOWNLOAD_DIR = BASE / "downloads"
DATA_DIR.mkdir(exist_ok=True)
DOWNLOAD_DIR.mkdir(exist_ok=True)

load_dotenv(BASE / ".env")
CONFIG = json.loads((BASE / "config.json").read_text(encoding="utf-8"))

MIS_URL = os.getenv("CEPT_MIS_URL", CONFIG.get("mis_url", ""))
USERNAME = os.getenv("CEPT_USERNAME", "")
PASSWORD = os.getenv("CEPT_PASSWORD", "")
AUTO_PUSH = os.getenv("GIT_AUTO_PUSH", "true").lower() == "true"

if not MIS_URL or not USERNAME or not PASSWORD:
    sys.exit("Please set CEPT_MIS_URL, CEPT_USERNAME and CEPT_PASSWORD in automation/.env")


def download_reports():
    """Login to CEPT MIS and download each configured BO-wise report."""
    downloaded = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=CONFIG.get("headless", True))
        page = browser.new_page(accept_downloads=True)
        page.goto(MIS_URL)
        page.fill(CONFIG["login"]["username_selector"], USERNAME)
        page.fill(CONFIG["login"]["password_selector"], PASSWORD)
        page.click(CONFIG["login"]["submit_selector"])
        page.wait_for_load_state("networkidle")

        for report in CONFIG["reports"]:
            page.goto(report["url"])
            for selector, value in report.get("form", {}).items():
                page.fill(selector, value)
            if report.get("submit_selector"):
                page.click(report["submit_selector"])
            page.wait_for_load_state("networkidle")
            with page.expect_download() as dl_info:
                page.click(report["download_selector"])
            target = DOWNLOAD_DIR / f"{report['name']}_{date.today().isoformat()}.xlsx"
            dl_info.value.save_as(str(target))
            downloaded.append((report["name"], target))
        browser.close()
    return downloaded


def consolidate(downloaded):
    """Merge all downloaded BO-wise reports into one dashboard-format Excel."""
    frames = []
    for name, path in downloaded:
        df = pd.read_excel(path)
        df.columns = [str(c).strip() for c in df.columns]
        date_col = next((c for c in df.columns if "date" in c.lower()), None)
        bo_col = next((c for c in df.columns if c.lower() in ("bo", "bo name", "office_name", "office name") or "branch" in c.lower()), None)
        if not date_col or not bo_col:
            print(f"Skipping {name}: Date/BO column not found. Columns: {list(df.columns)}")
            continue
        df = df.rename(columns={date_col: "Date", bo_col: "BO"})
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        df["BO"] = df["BO"].astype(str).str.strip()
        keep = ["Date", "BO"] + [c for c in df.columns if c not in ("Date", "BO") and pd.api.types.is_numeric_dtype(df[c])]
        frames.append(df[keep])
    if not frames:
        sys.exit("No valid BO-wise data found in downloaded reports.")
    out = pd.concat(frames, ignore_index=True)
    out = out.groupby(["Date", "BO"], as_index=False).sum(numeric_only=True)
    target = DATA_DIR / "dashboard_data.xlsx"
    out.to_excel(target, index=False)
    print(f"Consolidated {len(out)} rows -> {target}")
    return target


def push_to_github():
    if not AUTO_PUSH:
        return
    subprocess.run(["git", "add", "data/dashboard_data.xlsx"], cwd=ROOT, check=False)
    subprocess.run(["git", "commit", "-m", f"Auto-update dashboard data {date.today().isoformat()}"], cwd=ROOT, check=False)
    subprocess.run(["git", "push"], cwd=ROOT, check=False)
    print("Pushed updated dashboard data to GitHub.")


if __name__ == "__main__":
    files = download_reports()
    consolidate(files)
    push_to_github()
