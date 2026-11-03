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
- **Auto-load**: `data/history.csv` (daily snapshots) + `data/exports/manifest.json` me listed KPI exports automatically load ho kar BO-wise table, KPI cards aur trend me pivot ho jate hain. Header me **"Data as of"** date dikhti hai. Manual upload agar pehle ho gaya to auto-load usse overwrite nahi karta.
- BO-wise cumulative performance table with totals:
  - **Click-to-sort** har column par (BO/SO/metric/Total)
  - **Total column** per row + totals footer
  - **Quick view**: Top 10 / Bottom 10 by total
  - **Group by toggle**: BO-wise ya SO-wise aggregation
- Cumulative daily trend chart (history file ki wajah se multi-day trend sach me dikhta hai; latest export us date ko replace karta hai taaki re-run double-count na kare)
- Flexible date filter:
  - **Upto date**: start date se kisi bhi selected date tak
  - **Date range**: kisi bhi from–to range ka data (From field sirf range mode me dikhta hai)
  - **Single date**: sirf ek din ka data
- BO-wise filter (All BOs ya koi ek BO) with master-list pre-set and fuzzy name matching
- **Poora BO master list**: BO-wise table me `bo_master.js` ke saare BOs hamesha dikhte hain. Jin BOs ne selected date window me koi account nahi khola unki row **0** ke saath aati hai (report se gayab hone se row drop nahi hoti). Excel export ka BO-wise sheet bhi isi list ko mirror karta hai.
- **Reporting period `2026-11-01` se**: is date se pehle ke snapshots (jaise 31-10-2026) KPI cards, trend, BO table aur Excel/PDF export me include nahi hote. Purani history audit ke liye `history.csv` me rah sakti hai, lekin dashboard sirf 01-11-2026 onwards dikhata hai.
- **SO filter** — Sub Office wise filter (B.O rows apne parent S.O ke saath aate hain)
- **Accounts Opened** per BO (product columns ka sum) + totals row footer me
- **Office summary rows kept separate**: CEPT report ki S.O/H.O/G.P.O summary rows (blank `SO Name` cell, ya `S.O`-ending name — master-list BO kabhi nahi) `kind='so'` rows ban kar rakhi jati hain. BO-wise table, KPI cards aur trend chart inhe kabhi include nahi karte (double-counting nahi). **SO-wise view me har SO ke liye do alag sortable columns**: `<metric> (ΣBO)` = us SO ke BO rows ka sum, aur `<metric> (SO)` = office ki apni summary row ka reported figure (`—` agar report me woh row nahi). `(no SO)` bucket = woh BO rows jinki SO Name blank hai (jaise `Brahman Majra B.O`)
- **KPI cards par day-over-day delta** (▲/▼ vs previous day) — history file ki wajah se
- **Export Excel**: filtered raw data + BO-wise pivot, dono sheets me
- **Export PDF**: browser print dialog se PDF save (print header me active filters + data-as-of date ke saath)
- **Security**: CSP meta (script-src: self + jsdelivr only — isliye koi inline `on*` handler nahi chalega) + SRI hashes on dono CDN scripts; saara JS `app.js` me hai aur saare event listeners wahin `addEventListener` se wire hote hain (inline script/handlers nahi)

## Auto-fed data files

| File | Written by | Format |
|---|---|---|
| `data/exports/{kpi}_merged.xlsx` | `cept_mis_cbs.py` | one header row + BO data rows |
| `data/exports/manifest.json` | `write_manifest()` | `[{"file": "data/exports/...", "label": "Accounts Opened", "date": "2026-11-01"}]` — dashboard isi list ko auto-load karta hai (naya KPI add karne ke liye frontend change nahi chahiye). `date` report ki apni date hai kyunki merged file me `Dated:` title row nahi rehta |
| `data/history.csv` | `append_history()` | long format: `date,bo,so,sol,kind,kpi,value` — `kind` = `bo` ya `so` (office summary row). Har report date ki rows us date ki purani rows ko replace karti hain (idempotent) |
| `data/dashboard_data.xlsx` | `generate_dashboard_data()` | one row per BO (Date/BO/SO/SOL ID + KPI columns) |

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
