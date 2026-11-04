# BO Dashboard

Daily Excel upload se **BO-wise cumulative performance** dikhane wala fast, single-page dashboard. Koi backend/database nahi — sab kuch browser me process hota hai, isliye filtering aur rendering fraction of seconds me hoti hai.

Deployment: GitHub Pages via GitHub Actions (re-triggered after repo made public).

## Excel format

File me minimum ye columns hone chahiye (header row ke saath):

| Column | Type | Example |
|---|---|---|
| `Date` | Date | 2026-10-01 |
| `BO` | Text | BO Ambala |
| Performance columns | Numeric | Accounts Opened, Deposits, Advances, Recovery |

Aap kitne bhi numeric columns rakh sakte hain — dashboard unhe automatically detect karke KPI cards, cumulative chart aur BO-wise table me dikha dega. Template ke liye `sample-data.csv` dekhen.

## Features

- Daily Excel (.xlsx/.xls) ya CSV upload — drag & drop supported
- **Auto-load**: `data/history.csv` (daily snapshots) + `data/exports/manifest.json` me listed KPI exports automatically load ho kar BO-wise table, KPI cards aur trend me pivot ho jate hain. Header me **"Data as of"** date dikhti hai. Manual upload agar pehle ho gaya to auto-load usse overwrite nahi karta. UI, print header, chart labels aur Excel export me dates **DD-MM-YYYY** (jaise `02-11-2026`) format me dikhti hain; internal/history data ISO (`YYYY-MM-DD`) hi rehta hai.
- BO-wise cumulative performance table with totals:
  - **Click-to-sort** har column par (BO/SO/metric/Total)
  - **Total column** per row + totals footer
  - **Quick view**: Top 10 / Bottom 10 by total
  - **Group by**: BO-wise / SO-wise / Division-wise / Sub-division-wise aggregation
- Cumulative daily trend chart (history file ki wajah se multi-day trend sach me dikhta hai; latest export us date ko replace karta hai taaki re-run double-count na kare)
- Flexible date filter:
  - **Upto date**: start date se kisi bhi selected date tak
  - **Date range**: kisi bhi from–to range ka data (From field sirf range mode me dikhta hai)
  - **Single date**: sirf ek din ka data
- **Office identity = SOL ID / BOCODE** (`bo_master.js` ka `code`). Naam se milane wala fuzzy matching hata diya gaya hai: master sheet me 2,196 offices hain lekin (canonical form me) sirf 2,093 alag naam — 82 naam 184 offices ke beech share ho rahe hain (`Aharwan BO` ek Faridabad, ek Hissar; `Alika BO` teen), isliye naam akela identity nahi ho sakta. Report ka code master code se match hua to wahi office; code se na milne par sirf unique name dekha jata hai (ambiguous name = koi match nahi).
- **Poora BO master list**: BO-wise table me `bo_master.js` ke saare 2,196 offices hamesha dikhte hain (report ke offices ∪ master sheet). Jin offices ne selected date window me koi account nahi khola unki row **0** ke saath aati hai (report se gayab hone se row drop nahi hoti). Excel export ka BO-wise sheet bhi isi list ko mirror karta hai.
- **Reporting period `2026-11-01` se**: is date se pehle ke snapshots (jaise 31-10-2026) KPI cards, trend, BO table aur Excel/PDF export me include nahi hote. Purani history audit ke liye `history.csv` me rah sakti hai, lekin dashboard sirf 01-11-2026 onwards dikhata hai.
- **BO filter / SO filter** — type-to-search (master sheet ke saare BO aur SO names options ke saath)
- **Division aur Sub Division filter** — `bo_master.js` ke `division`/`subDivision` columns se, jo uploaded `BO_Master_data_of_Haryana_Circle.xlsx` (sheet `Consolidate`) se bane: 9 divisions, 33 sub-divisions (ek Hissar office ka sub-division sheet me khali hai, woh `(no sub division)` bucket me aata hai). Division select karne par sub-division list us division tak narrow ho jata hai, aur filter AND hota hai. Division-wise / Sub-division-wise group-by me master sheet ki har division/sub-division ki row milti hai (jisme is window me kuch nahi khula uska figure 0), KPI cards par Divisions/Sub Divisions counts dikhte hain, aur Excel export me `Division-wise` + `SubDivision-wise` sheets banti hain.
- **Office rows ka division**: SO/H.O summary rows ke apne koi division nahi hote, isliye woh master sheet ke SO column (aur zarurat par BO rows) se inherit karte hain — jis SO ke BOs ek se zyada sub-division me faile hue hain, uska figure sirf division level par dikhta hai (`(no sub division)` bucket me).
- **Accounts Opened** per BO (product columns ka sum) + totals row footer me
- **BO tier counts — `0 Accounts` / `Under 10 Accounts` / `Under 100 Accounts`** KPI cards. Teeno counts exactly usi scope par gate jate hain jis par **BOs** card ginta hai: koi division/sub-division select nahi kiya to poore circle ke BOs (overall), division select kiya to us division ke BOs, sub-division select kiya to us sub-division ke BOs (BO/SO type-to-search laga ho to uske baad bache offices). Count cumulative hai — `Under 10` me 0-wale BOs bhi shamil, `Under 100` me 0 aur under-10 dono; har card ke note me scope aur "of N BOs in scope" likha hota hai. Per-BO figure = selected date window me us office ke accounts ka sum (metric cards me `Accounts` wala metric select hota hai; na milne par pehla metric, aur note uska naam dikhata hai).
- **Office summary rows kept separate**: CEPT report ki S.O/H.O/G.P.O summary rows (blank `SO Name` cell, ya name me `S.O`/`H.O`/`P.O`/`G.P.O` marker, ya 8-digit all-numeric code — master-list BO kabhi nahi) `kind='so'` rows ban kar rakhi jati hain. `history.csv` ki `kind` column se zyada master sheet ki maan hoti hai: code master BO ka hai to row `bo` hi rahegi. BO-wise table, KPI cards aur trend chart inhe kabhi include nahi karte (double-counting nahi). **SO-wise view me har SO ke liye do alag sortable columns**: `<metric> (ΣBO)` = us SO ke BO rows ka sum, aur `<metric> (SO)` = office ki apni summary row ka reported figure (`—` agar report me woh row nahi). `(no SO)` bucket = woh BO rows jinki SO Name blank hai (jaise `Brahman Majra B.O`)
- **KPI cards par day-over-day delta** (▲/▼ vs previous day) — history file ki wajah se
- **Export Excel**: 4 sheets — `Filtered Data` (raw rows + Division/Sub Division), `BO-wise` (har master office, `BO ID` code ke saath), `Division-wise` aur `SubDivision-wise` roll-ups (total ke order me)
- **Export PDF**: browser print dialog se PDF save (print header me active filters + data-as-of date ke saath)
- **Bright (light) theme**: soft off-white background, white cards with a gentle shadow, dark slate ink and a calm teal accent — comfortable for long reading, with zebra-striped table rows, sticky header/totals row and a focus ring on inputs. Chart lines use a dark-on-white palette (legend/ticks grey) so the trend stays readable. Print/PDF export forces plain black-on-white regardless of the theme. `tests/dashboard.test.js` asserts WCAG-grade contrast (background/card near white, ink >= 7:1, muted text and accent >= 4.5:1, chart colours >= 4.5:1 on white), so a dark palette cannot creep back in
- **Security**: CSP meta (script-src: self + jsdelivr only — isliye koi inline `on*` handler nahi chalega) + SRI hashes on dono CDN scripts; saara JS `app.js` me hai aur saare event listeners wahin `addEventListener` se wire hote hain (inline script/handlers nahi)

## Auto-fed data files

| File | Written by | Format |
|---|---|---|
| `data/exports/{kpi}_merged.xlsx` | `cept_mis_cbs.py` | one header row + BO data rows |
| `data/exports/manifest.json` | `write_manifest()` | `[{"file": "data/exports/...", "label": "Accounts Opened", "date": "2026-11-01"}]` — dashboard isi list ko auto-load karta hai (naya KPI add karne ke liye frontend change nahi chahiye). `date` report ki apni date hai kyunki merged file me `Dated:` title row nahi rehta |
| `data/history.csv` | `append_history()` | long format: `date,bo,so,sol,kind,kpi,value` — `kind` = `bo` ya `so` (office summary row). Har report date ki rows us date ki purani rows ko replace karti hain (idempotent) |
| `data/dashboard_data.xlsx` | `generate_dashboard_data()` | one row per BO (Date/BO/SO/SOL ID + KPI columns) |

## BO master file (`bo_master.js`)

`const BO_MASTER = [{name, code, so, subDivision, division}, ...]` — 2,196 Haryana Circle offices,
generated from `BO_Master_data_of_Haryana_Circle.xlsx` (sheet `Consolidate`): `Office_Name`,
`SOL_ID_BO_ID` (upper-cased; `0`/blank codes stored as empty), `AO_Name` → SO, `Sub_Division`,
`Division`. File ek valid JSON array hai, isliye `cept_mis_cbs.py` aur dashboard dono isi ko padhte
hain. Dashboard ke liye yahi file identity (code), SO/Division/Sub Division hierarchy aur "0 wale BO"
ki list deti hai — sheet me naya office add karne par is file ko regenerate karke commit karna hota hai.

## Local run

```bash
python -m http.server 8000
# phir browser me: http://localhost:8000

# smoke test (local ya live URL):
DASHBOARD_URL=http://localhost:8000/ python -u automation/smoke_test.py
```

## GitHub Pages & CI

Repo public hai aur `.github/workflows/pages.yml` har push par site ko deploy karta hai.

Live URL: https://vcharyanaco-tech.github.io/BO-dashboard/

Workflows:

- `cept-mis.yml` — **single** automation entry point (pehle `cept-mis.yml` + `dashboard.yml` dono chal rahe the aur main par race karte the; ab ek hi hai, din me 2 baar cron + concurrency group ke saath). Run ke baad `automation/sanity_check.py` naye merged exports ko HEAD se compare karta hai — agar BO rows 50%+ gir gaye ya file khali hai to commit nahi hota. Commit karta hai: merged exports, `manifest.json`, `history.csv`, `dashboard_data.xlsx`. Data commit ke baad workflow `pages.yml` ko dispatch karta hai, kyunki `GITHUB_TOKEN` se kiya gaya push apne aap dusre workflows trigger nahi karta — warna live dashboard purane data par atka rehta hai.
- `test.yml` — har push/PR par `node tests/dashboard.test.js` (dashboard logic + repo data files) aur Python syntax check chalata hai.
- `smoke-test.yml` — Pages deploy ke baad (ya manually) live URL ko headless Chromium me load karke check karta hai ki KPI cards + BO table render hue aur koi page error nahi aaya.
- `backfill_dashboard.yml`, `debug_merged_files.yml` — manual helpers.
