# CEPT MIS Automation

CEPT MIS ke **CBSTrans3** report page (`https://mis.cept.gov.in/CBS/CBSTrans3.aspx`) se
BO-wise data nikaal kar dashboard ke liye Excel files banata hai.

## Primary script: `cept_mis_cbs.py`

Flow (per division/sub-division, per KPI):

1. `CBSTrans3.aspx` khulta hai, reporting date `DD-MM-YYYY` me bharte hain.
2. Cascade dropdowns: Circle → Region → Division → Sub Division. Har step ke baad
   dropdown me placeholder ("Select ...") ke alawa asli option aane ka wait hota hai,
   isliye postback ke beech me select karne se option miss nahi hota.
3. KPI radio select karke submit, phir report render hone ka wait. Wait iframe-aware hai
   (`Dated:` header ya `SOL ID / BOCODE` table kisi bhi frame me mile to ready).
4. `Dated:`/table milne par DataTables Excel control se file download hoti hai.

KPI: `accounts_opened` (`#ContentPlaceHolder1_ACO`), `accounts_closed`
(`#ContentPlaceHolder1_ACC`), `net_addition` (`#ContentPlaceHolder1_NET`).

### Output

- `data/exports/{kpi}_{division}_{subdivision}.xlsx` — per sub-division intermediate (run ke ant me delete)
- `data/exports/{kpi}_merged.xlsx` — ek KPI ke saare rows ek file me (3 files)
- `debug/*.png` + `debug/*.html` — failure par page snapshot (CI me artifact ke roop me upload hota hai)

## Setup

```bash
cd automation
pip install -r requirements.txt
playwright install --with-deps chromium
```

## Run

```bash
python automation/cept_mis_cbs.py
```

CI (`.github/workflows/cept-mis.yml`) isi script ko daily chalata hai aur
`data/exports/*_merged.xlsx` commit karta hai. Chromium container/GitHub Actions me
`--no-sandbox --disable-dev-shm-usage` ke saath launch hota hai.

## Legacy script: `cept_mis_downloader.py`

Login-based purana flow (`.env` me credentials, `config.json` me report URL). Ye
CBSTrans3 ke naye Excel-export flow se replace ho chuka hai; sirf reference ke liye rakha hai.

## Important

- Credentials kabhi repo me commit na karein; local `.env` hi use karein.
- Selectors MIS UI ke saath badalte hain — change hone par `cept_mis_cbs.py` ke top par
  diye gaye `*_SEL` constants update karein.
- Page ka behaviour verify karne ke liye `headless=True` rakhein; debugging ke liye
  `debug/` folder dekhein.
