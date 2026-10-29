# CEPT MIS Auto-Downloader

Ye tool aapke system par chal kar CEPT MIS se BO-wise reports download karta hai, unhe ek single `data/dashboard_data.xlsx` me consolidate karta hai, aur (optional) GitHub par push karta hai jise dashboard automatically load kar leta hai.

## Setup

```bash
cd automation
pip install -r requirements.txt
playwright install chromium
copy .env.example .env   # Windows
# .env me apne CEPT credentials bharein
```

## Run

```bash
python cept_mis_downloader.py
```

## Daily automation (Windows Task Scheduler)

1. Task Scheduler → Create Basic Task
2. Trigger: Daily (e.g. 6:00 PM)
3. Action: Start a program → `python` → arguments: `path\to\automation\cept_mis_downloader.py`

## Important

- CEPT MIS ka login URL aur report-page selectors `automation/config.json` me configurable hain — MIS UI change hone par sirf wahan update karein.
- Credentials sirf local `.env` me rakhen; use kabhi GitHub par push na karein.
- Agar CEPT MIS sirf department network/intranet se accessible hai, to ye script usi network/system par chalani hogi.
- Pehli baar chalate waqt `headless: false` rakh kar selectors verify karein, phir `true` kar dein.