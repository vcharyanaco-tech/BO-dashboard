# CEPT MIS Automation

CEPT MIS ke **CBSTrans3** report page (`https://mis.cept.gov.in/CBS/CBSTrans3.aspx`) se
BO-wise data nikaal kar dashboard ke liye Excel files banata hai.

## Primary script: `cept_mis_cbs.py`

Flow (per division/sub-division, per KPI) — live page par verify kiya gaya:

1. `CBSTrans3.aspx` khulta hai. Ye classic ASP.NET WebForms page hai
   (`__VIEWSTATE` / `__EVENTVALIDATION`) — har filter change ek full `__doPostBack` hai.
2. Reporting dates do text inputs hain: `startDate` aur `endDate`, jquery-ui
   datepicker ke saath, format **`DD-MM-YYYY`** (`dateFormat: "dd-mm-yy"`).
3. Page header par **"Data Updated upto: DD/MM/YYYY"** label hota hai. Us date se
   aage ki window **header-only (0 BO rows)** report deti hai, isliye window hamesha
   usi date tak clamp hoti hai.
4. Cascade dropdowns: Circle → Region → Division → Sub Division. Har step ke baad
   dropdown me placeholder ("Select ...") ke alawa asli option aane ka wait hota hai,
   isliye postback ke beech me select karne se option miss nahi hota.
5. **Submit** button (`button.btn-danger`, `__doPostBack('ctl00$ContentPlaceHolder1$ctl00','')`)
   current filters + dates ke saath report render karta hai. Grid `#example1` hai.
6. Sirf Accounts Opened radio (`#ContentPlaceHolder1_ACO`, name `Group1`) tick hota hai
   aur uska postback grid ko re-render karta hai. Wait real BO rows par hota hai —
   sirf header aana kaafi nahi.
7. Grid ready hone par DataTables ka **Excel button** click hota hai.

### Export (client-side, koi server export URL nahi)

Grid `dom: 'Bifrtp'` ke saath initialize hota hai aur buttons me
`{ extend: 'excel', footer: true }` hai — yaani workbook **browser me** banti hai aur
button `#example1_wrapper .dt-buttons button.buttons-excel` par hota hai.

- Page par teen DataTables hain (`#example`, `#example1`, `#example3`), isliye
  selector `#example1_wrapper` ke andar scope kiya gaya hai — warna kisi doosri
  table ka Excel button click ho sakta hai.
- Filename `document.title` se aata hai (`Product Wise A/C Report`), isliye script
  download ko `download.suggested_filename` ke bajaye apne target path par save karti hai.

### Export shape (live page se confirm kiya)

`#example1` ka header row: `Name`, `SOL ID / BOCODE`, `SO Name`, phir 12 product columns
(`MIS`, `PPFGP`, `SSA`, `RD`, `SBBAS`, `SBSGP`, `SCSS`, `TD`, `PRFTS`, `KVN`, `NSC8`, `MSSC`).
**Koi `Total` data column nahi hai** — ek BO ka KPI value in product cells ka sum hai.

Data rows seedhe `<tr>` hote hain — table me **`<tbody>` nahi hai**. `<tfoot>` me ek
`Total:` row hoti hai jise DataTables bharta hai aur `footer: true` ki wajah se Excel
export me bhi likhta hai; `merge_kpi_files` aur `generate_dashboard_data` usse hata dete
hain, warna dashboard use ek BO samajh leta hai.

### BO master list (`bo_master.js`)

`python3 automation/gen_bo_master.py <BO_Master_data_of_Haryana_Circle.xlsx>` repo root me
`bo_master.js` banata hai — `Consolidate` sheet ke 2,196 offices as
`[{name, code, so, subDivision, division}, ...]` (valid JSON array, isliye yehi file
`cept_mis_cbs.py` bhi padhta hai). `code` = `SOL_ID_BO_ID` hi identity hai: 2,093 hi distinct
naam hain, 82 naam 184 offices me share ho rahe hain, isliye naam se fuzzy match nahi hota.
`is_master_office()` / `is_so_row()` isi file se codes aur naam check karte hain.

### Output

- `data/exports/{kpi}_{division}_{subdivision}.xlsx` — per sub-division intermediate (run ke ant me delete)
- `data/exports/{kpi}_merged.xlsx` — har KPI ke saare rows, ek header row + sirf data rows. Office summary rows (`Adhoya S.O`, `Jamalpur SO Bhiwani`, `Ambala G.P.O.` — `is_so_row()`: master BO code/naam nahi, aur name me S.O/H.O/P.O/G.P.O marker ya 8-digit all-numeric code) `kind='so'` ke saath likhi jati hain, taaki dashboard me woh BO totals me na ghusen
- `data/exports/manifest.json` — dashboard ko batata hai kaun-kaun se merged files auto-load karne hain, har entry me report `date` bhi (merged file me `Dated:` title row nahi rehta). Naya KPI `METRICS` me add karne par dashboard me koi change nahi chahiye
- `data/history.csv` — daily snapshots, long format `date,bo,so,sol,kind,kpi,value` where `kind` = `bo` | `so` (office summary row: blank SO Name cell, office-style name ya all-digit code, and never a master BO code/naam). Dashboard `kind` se zyada `sol` code ki maan karta hai — code master BO ka hua to row BO hi rahegi. Same date ka re-run us date ki rows replace karta hai (idempotent). Isi se dashboard ka cumulative trend chart multi-day hota hai aur KPI cards ka day-over-day delta kaam karta hai. `kind='so'` rows SO-wise view me ΣBO se alag dikhte hain. ~150 KB/day grow karta hai — bahut purana data prune karne ke liye `append_history()` me ek retention window add ki ja sakti hai
- `data/dashboard_data.xlsx` — dashboard ka BO-wise file (report window ki end date ke saath)
- `debug/*.png` + `debug/*.html` — failure par page snapshot (CI me artifact ke roop me upload hota hai)

CI (`cept-mis.yml`) run ke baad `automation/sanity_check.py` naye merged files ko git HEAD wali copy se compare karta hai — BO rows 50%+ kam ya khali file milne par commit nahi hota. Deploy ke baad `smoke-test.yml` live page ko check karta hai.

## Date window

Default: kal ki single-day window (jaise pehle tha), lekin **"Data Updated upto" se
aage kabhi nahi**. Override karne ke liye:

```bash
REPORT_START_DATE=01-10-2026 REPORT_END_DATE=30-10-2026 python automation/cept_mis_cbs.py
```

`REPORT_CIRCLE` (default `Haryana Circle`) aur `REPORT_REGION` (default
`DivReportingCircle`) se cascade bhi override ho sakta hai.

## Report granularity

`REPORT_LEVEL=division` exports one workbook per division per KPI instead of one per
sub-division (default). Division-level returns the same BO rows - verified on Ambala
Division: 308 ACO rows at division level = sum of its four sub-divisions (81 + 72 + ...) -
with ~4x fewer page loads, which matters because the server slows down and can answer
HTTP 500 under sustained load.

## Dashboard columns

`data/dashboard_data.xlsx` carries `Date | BO | SO | SOL ID | Accounts Opened`.

The dashboard reports one KPI. `Accounts Opened` is the per-BO sum of the product cells
(`MIS`, `PPFGP`, `SSA`, ...), and the totals row at the foot of the table sums the BOs
currently shown. `SO` is the parent Sub Office of a B.O row; it is blank for S.O rows
because they *are* the Sub Office (the source grid does not print a parent for them).

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
  diye gaye `*_SEL` / `EXPORT_SELECTORS` constants update karein.
- Page ka behaviour verify karne ke liye `headless=True` rakhein; debugging ke liye
  `debug/` folder dekhein.
