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
6. KPI radio (`#ContentPlaceHolder1_ACO` / `_ACC` / `_NET`, name `Group1`) har ek
   postback karke `#example1` ko us KPI ke liye re-render karta hai. Wait real BO rows
   par hota hai — sirf header aana kaafi nahi.
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

### Output

- `data/exports/{kpi}_{division}_{subdivision}.xlsx` — per sub-division intermediate (run ke ant me delete)
- `data/exports/{kpi}_merged.xlsx` — ek KPI ke saare rows, ek header row + sirf data rows (3 files)
- `data/dashboard_data.xlsx` — dashboard ka BO-wise file
- `debug/*.png` + `debug/*.html` — failure par page snapshot (CI me artifact ke roop me upload hota hai)

## Date window

Default: kal ki single-day window (jaise pehle tha), lekin **"Data Updated upto" se
aage kabhi nahi**. Override karne ke liye:

```bash
REPORT_START_DATE=01-10-2026 REPORT_END_DATE=30-10-2026 python automation/cept_mis_cbs.py
```

`REPORT_CIRCLE` (default `Haryana Circle`) aur `REPORT_REGION` (default
`DivReportingCircle`) se cascade bhi override ho sakta hai.

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
