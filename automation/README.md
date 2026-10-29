# CEPT MIS Automation

CEPT MIS ke **CBSTrans3** report page (`https://mis.cept.gov.in/CBS/CBSTrans3.aspx`) se
BO-wise data nikaal kar dashboard ke liye Excel files banata hai.

## Primary script: `cept_mis_cbs.py`

Flow (per division/sub-division, per KPI):

1. `CBSTrans3.aspx` khulta hai, reporting date `DD-MM-YYYY` me bharte hain.
2. Cascade dropdowns: Circle → Region → Division → Sub Division. Har step ke baad
   dropdown me placeholder ("Select ...") ke alawa asli option aane ka wait hota hai,
   isliye postback ke beech me select karne se option miss nahi hota.
3. KPI radio select karne par report render hota hai. Wait `#example1` grid me koi asli BO
   row (jo `Total:` nahi hai) aane tak chalta hai, kisi bhi frame me. Sirf `Dated:` header
   kaafi nahi hai — filters khali hon to bhi header print ho jata hai.
4. Grid ready hone par DataTables ka **Excel button** (`.dt-buttons button.buttons-excel`)
   click hota hai. Workbook browser me banti hai (server-side export page nahi hai) aur
   download apne `{kpi}_{division}_{subdivision}.xlsx` naam se save hoti hai.

KPI: `accounts_opened` (`#ContentPlaceHolder1_ACO`), `accounts_closed`
(`#ContentPlaceHolder1_ACC`), `net_addition` (`#ContentPlaceHolder1_NET`).

### Output

- `data/exports/{kpi}_{division}_{subdivision}.xlsx` — per sub-division intermediate (run ke ant me delete)
- `data/exports/{kpi}_merged.xlsx` — ek KPI ke saare rows ek file me (3 files)
- `debug/*.png` + `debug/*.html` — failure par page snapshot (CI me artifact ke roop me upload hota hai)

## Export shape (live page se confirm kiya)

`#example1` ka header row: `Name`, `SOL ID / BOCODE`, `SO Name`, phir 12 product columns
(`MIS`, `PPFGP`, `SSA`, `RD`, `SBBAS`, `SBSGP`, `SCSS`, `TD`, `PRFTS`, `KVN`, `NSC8`, `MSSC`).
**Koi `Total` column nahi hai** — ek BO ka KPI value in product cells ka sum hai.

Grid ke `<tfoot>` me ek `Total:` row hoti hai aur Excel button `footer: true` ke saath usse
export me bhi likhta hai. `merge_kpi_files` usse hata deta hai, warna dashboard use ek BO
samajh leta hai.

Excel button filename `document.title` se leta hai, isliye har division/sub-division ke liye
download ka suggested naam same aata hai (`Product Wise A/C Report.xlsx`). Isliye script
download ko `download.suggested_filename` ke bajaye apne target path par save karti hai.

Date window: page par "Data Updated upto" date hoti hai; usse aage ki range khali grid deti
hai (0 BO rows) chahe export ho jaye. Isliye range hamesha available data ke andar rakhein.

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
