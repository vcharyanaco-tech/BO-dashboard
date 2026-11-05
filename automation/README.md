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

CI (`cept-mis.yml`) run ke baad `automation/sanity_check.py` gate lagata hai. Gate **coverage** ka hai, volume ka nahi: supervisor `debug/coverage.json` me likhta hai kaun-kaun se sub divisions ne download kiya, aur ek bhi planned sub division *answer hi na kare* to commit nahi hota. Zero-row jawab covered hai: report render hui par BO rows 0 thi to child `debug/empty_reports.json` me receipt chhodta hai aur wo sub division missing nahi geli jati, retry bhi nahi hoti. Operator override do tarah se: `SANITY_ALLOW_MISSING_SUBDIVISIONS=<n>` env, ya dispatch ka `allow_missing_subdivisions` input (MIS ke permanently wedged sub divisions ke liye). Kam rows wala din normal hai — kai din sirf kuch BOs account kholte hain — isliye 50%-drop rule ab sirf tab fail karta hai jab coverage report hi na ho (yanim merge supervisor ke bahar bana), warna woh WARN hai. Row count se ek cheez pakdi nahi jati: file byte-identical to HEAD ho (wahi report dobara aa gayi) to woh WARN karta hai, kyunki us date ka "snapshot" asal me pichle date ki nakal hota hai. Deploy ke baad `smoke-test.yml` live page ko check karta hai.

## Sunday = chhutti, report nahi

Hara office Sunday ko band rehta hai, isliye **Sunday ki report kabhi fetch nahi hoti**:

- `resolve_window()` ka default "kal" agar Sunday nikle to ek din peeche **Saturday** ho
  jaata hai — warna Monday ke run poora Saturday chhod dete aur history me ek din ka gap
  pad jaata. (Saturday ki dobara download hone par "ek snapshot = ek din" wala rule use
  doosre date par file nahi karta.)
- Phir bhi window sirf-ek-Sunday ho (`start == end` aur Sunday) to sweep bilkul nahi chalta:
  `sweep_and_build()` log karke bina download ke bahar aa jaata hai, aur supervisor bache
  hue children dispatch hi nahi karta. Run green rehta hai, commit kuch nahi hota.
- `REPORT_START_DATE`/`REPORT_END_DATE` se maangi gayi Sunday **silent rollback nahi** karti
  — use refusal ka reason log me dikhta hai, taaki operator ko pata rahe ki MIS ne data
  diya ya chhutti thi. Range jisme koi working day pade (Sat→Sun) wo fetch hoti hai.

## Date window

Default: kal ki single-day window (jaise pehle tha), lekin **"Data Updated upto" se
aage kabhi nahi**. Override karne ke liye:

```bash
REPORT_START_DATE=01-10-2026 REPORT_END_DATE=30-10-2026 python automation/cept_mis_cbs.py
```

Wahi window CI se bhi maang sakte hain: `cept-mis.yml` ke **Run workflow** me
`report_start_date` / `report_end_date` (DD-MM-YYYY) inputs hain, jo usi step ke
`REPORT_START_DATE` / `REPORT_END_DATE` env vars ban jate hain. Dono khali chhoden to
 behaviour pehle jaisa — scheduled/push runs hamesha kal ki (UTC) single-day window
lete hain, aur "Data Updated upto" se aage kabhi nahi. Ek din ki ghalti sudharne ke
liye start = end karke dispatch karein; us date ki history rows replace ho jayengi.

`REPORT_CIRCLE` (default `Haryana Circle`) aur `REPORT_REGION` (default
`DivReportingCircle`) se cascade bhi override ho sakta hai.

## Report granularity

`REPORT_LEVEL=division` (default `subdivision`) exports one workbook per division per
KPI. **Do not use it in CI**: on 04-11 the division-level page rendered 98 BO rows for
Ambala Division while its sub-divisions together reported ~300 (Bhiwani 71, Faridabad 56
against 431 and ~250 master offices), i.e. it silently drops most offices. The older
"same rows" note here was wrong. Sub-division level is what `cept-mis.yml` runs, and a
completed sweep takes ~20-25 minutes.

## Ek sub division = ek process (supervisor)

`main()` ab default supervisor hai: pehle ek bounded page visit (`CEPT_LIST=1`) se MIS ke
apne division **aur sub division** labels padhe jaate hain (`debug/divisions.json`), phir
**har sub division ke liye alag child process** chalta hai
(`CEPT_CHILD=1 SKIP_MERGE=1 REPORT_DIVISIONS=<division> REPORT_SUBDIVISIONS=<sub>`) jise
`SUB_BUDGET_S` (default 180s) ka hard budget milta hai. Budget paar kiya to process group
SIGKILL se mara jata hai, bache hue downloads disk par reh jaate hain, aur supervisor
baaki sub divisions chala kar jo mila use merge/build karta hai. jis division ka sub
division list page na de (visit fail hui), wo ek poora division child ban kar chalta hai
(`DIVISION_BUDGET_S`, default 600s) — aur log me ye clearly likha aata hai.

Kyun: `page.evaluate()` ka koi timeout nahi hota, isliye ek wedged MIS page poore sweep ko
rok deta hai — 04-11 ko Bhiwani/Mohindergarh ne 39 minute aur phir poora 100-minute run
kha liya, aur kyunki merge script ke aakh me hota hai, killed run **kuch commit nahi karta
tha**. Division-level splitting (commit 1029207) se din to bach gaya, par nuksan bada raha:
ek hi hang ne Bhiwani ke 4 aur sub divisions aur Gurgaon ka Narnaul (2 download ho chuke
the) sab kha liye — 42 me se 30 minute teen hangs par gaye aur din me sirf 409 of ~2,100
BO rows aayi, jise tab ka volume-gate rok baitha (galat gate — kam rows normal ho sakte
hain, chhoota hua sub division nahi). Ab hang ka cost ek sub division hai, aur gate
coverage hai.

Covered ka matlab "the page answered": har planned child ke baad supervisor disk par uska
workbook dhoondta hai (`got_data()` / `division_has_data()`), exit code se zyada usi par
bharosa hai. Report render hui par us din koi BO row nahi thi to child
`debug/empty_reports.json` me zero-row receipt chhodta hai — wo sub division **covered**
hai, retry nahi mangti (kam rows wala din normal hai). Jo child bina jawab ke laut aaya
(killed, failed, ya `select_option` timeout) use supervisor **`SUB_RETRIES` (default 2)
baar dobara bhejta hai**, is baar `RETRY_BUDGET_S` (default 360s) budget dekar — jab tak
poori sweep `SUPERVISOR_BUDGET_S` (default 7200s) ke andar hai. Phir bhi jo na aaye woh
`debug/coverage.json` me naam ke saath (`missing`) likha jaata hai aur sanity usi din ka
commit rok deti hai.

Aakh ke baad log se seedha pata chalta hai kaunsa bacha: `X/Y children filed data,
Z answered with 0 rows, N never answered`, phir har missing ka naam `no data for
<division> / <sub> (killed)` aur `no rows from <division>` (poori division tab, jab uska
ek bhi sub division file na kare — _partial_ division missing nahi bolta).

- `debug/divisions.json` / `debug/window.json` / `debug/coverage.json` /
  `debug/empty_reports.json` gitignored hain (sirf run ke andar ka kaam).
- Page labels zaroori hain: MIS "Sonepat Division" kehta hai, master sheet "Sonipat
  Division" — isliye supervisor page se hi naam leta hai; page na de to master sheet par
  fallback hota hai aur log me warning aati hai.
- `CEPT_INLINE=1` purana single-process sweep wapas laata hai (debug ke liye).
- `REPORT_DIVISIONS=rohtak,karnal` supervisor ki plan ko filter karta hai;
  `REPORT_SUBDIVISIONS=<label>` child ko ek hi sub division par band karta hai.
- Worst case time: 420s (label visit) + ~31 sub divisions × 180s, plus retries, capped by
  `SUPERVISOR_BUDGET_S` (120 min) — isliye workflow step `timeout 8400s` (140 min) ke andar
  rehta hai aur job 180 minute par.
- Cap bachhe hue waqt me lagta hai, khaaye hue me nahi: supervisor har child se pehle
  `in_time_for()` poochta hai — agar us child poora budget sweep ceiling ke andar nahi
  baithta to wo child **shuru hi nahi hota** aur log me `never started` / `not retried this
  round` ke saath naam aa jaata hai, `coverage.json` me exit `not started` ke saath. Pehle
  pehla pass bina guard tha aur retry sirf rounds ke *beech* rukta tha, isliye ek child ya
  chal rahi round 8,400s step timeout cross kar jaati thi aur run commit ke bina mar jaata.
- 04-11 ke do runs ka poora din 27/31 sub divisions par ruka tha (teen hang + ek khaali);
  ab wahi kaam ~13 minute me ho jaata hai aur baaki 4 retry rounds me aa jaate hain.
- 05-11 ka run (window 03-11, kyunki page abhi bhi "Data Updated upto: 03-11-2026" kehta
  hai) bhi 27/31 par ruka: Mohindergarh, Narnaul aur Sirsa-I har attempt par 360s wedge
  karte hain (MIS side), Bahadurgarh ka `select_option` 60s me timeout hota hai. Gate ne
  yahi 4 naam se commit roka — jo behaviour chahiye tha. Aise din ship karne ke liye
  dispatch me `allow_missing_subdivisions: 4` bharo.
- Locked in `tests/automation.test.py`: "one child per sub division, not per division",
  "a sub division that came back empty is retried until it reports", "a wedged sub division
  is retried, then named - never left silent", "coverage.json names exactly which sub
  divisions are missing", "with no sub division plan, one child sweeps a whole division",
  "a sweep that runs out of budget stops starting children, it does not overrun the job",
  "a sub division that reported zero rows is covered, not missing",
  "a zero-row sub division is not retried as a failed download",
  "no report is fetched for a Sunday, and no day is committed",
  "Monday's default window rolls back to Saturday instead of fetching a Sunday",
  "a fully covered day passes the gate even when few BOs opened accounts",
  "run_bounded kills a child that overspends its budget".

## Server errors

Mis kabhi kabhar report page ki jagah ASP.NET ka "Server Error in '/' Application —
Timeout expired … pooled connections" page bhej deta hai (DB connection pool exhausted).
Us page par filter dropdowns hote hi nahi, isliye pehle `select_option(CIRCLE_SEL)`
60s wait karke poora run maar deta tha. `open_page()` ab error page pehchan kar
(`page_is_error()`) max 4 baar reload karta hai (20s/40s/60s backoff) aur har baar
`debug/server_error_N.png|html` save karta hai; tab bhi fail hua to message ke saath
run fail hota hai, chup-chaap purana data commit nahi hota.

## Postback ke beech ka evaluate (select_radio retry)

Submit ka postback page ko dobara navigate karta hai; us navigation ke beech me
`page.evaluate()` mara to Playwright `Execution context was destroyed, most likely
because of a navigation` deta hai. `select_radio()` pehle us error ko retry nahi karta
tha, isliye 03-11-2026 wale run me 30 me se 27 sub divisions 2-16 second me hi haar gayi
thin — jabki usi waqt ka `debug/retry_*.html` report ka render hua page dikhata hai
(circle/division/sub division selected, `Dated: 03-11-2026 - 03-11-2026`, `#example1` me
BO rows). Ab `select_radio()` page settle hone tak dobara dispatch karta hai
(`settle_ms`, default 60s, beech me `domcontentloaded` wait) aur sirf window khatam hone
par error deta hai. Locked in `tests/automation.test.py` ("a navigation mid-select_radio
is retried, not fatal", "a page that never settles still raises").

## Ek office = ek row (identity keys)

History aur dashboard files me har office apne **code** se pehchana jata hai, naam se
nahi. `row_identity(kind, bo, sol)` wahi rule hai jo `app.js` lagata hai: `sol` me
`SOL_ID_BO_ID`/BOCODE hua to woh (uppercase, spaces hataakar) identity ban jata hai,
warna canonical naam, aur office summary rows `so:<name>` namespace me rehte hain (taaki
`Balu B.O` naam ke do alag offices kabhi ek row me na jud jayein — master sheet me aise
82 naam hain jo 184 offices me share ho rahe hain). Pehle `_read_merged()`/
`append_history()` naam se key karte the, jis se 01/11/2026 jaise history-served din me
45 per-BO figure galat aur ~96 offices missing/inflated the; ab code-keyed hai aur yeh
behaviour `tests/automation.test.py` me locked hai (`python3 tests/automation.test.py`).

## Ek snapshot = ek din

MIS ke kuch reports din bhar badalte nahi (same figures, dobara download ho jaate hain).
`append_history()` har date ka snapshot fingerprint (`office, kpi, value` ka sorted tuple)
nikalta hai: do alag dates ka fingerprint same hua to woh ek hi report ki do filing hai,
isliye **sirf newest stamp rakha jata hai** aur purani date history se gir jaati hai — log
me line aati hai `history: 2026-11-04 holds the same snapshot as 2026-11-05 - filed once,
as 2026-11-05`. Warn `sanity_check.py` se pehle hi milta hai (byte-identical workbook).
Is rule se pehle 31-10/01-11/02-11 teen din ban kar baithe the, aur cumulative card un
teeno ka jod (yaani ek hi report ka multiple) dikha raha tha. Locked in
`tests/automation.test.py` ("an identical snapshot is not filed as an extra day",
"the duplicate keeps the newer stamp").

## Download fail hua to run fail

Kisi bhi date ke liye ek bhi KPI workbook download nahi hui (server down, export grid
khali, "no valid BO data exported") to `main()` turant `sys.exit()` karta hai:
`Refusing to rewrite history/dashboard/manifest from a stale merged file`. Isse pehle
aise runs pichle run ki merged file usi raat ek naye date ke naam file kar dete the —
sanity check (row count) use pakad nahi sakta tha, aur history me 31-10/01-11/02-11 ek
hi snapshot ki teen nakal ban gayi thi (dashboard ka cumulative 2×40,087 = 80,174). Ab
koi download, koi commit.

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
