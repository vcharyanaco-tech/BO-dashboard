# BO Dashboard

Daily Excel upload se **BO-wise cumulative performance** dikhane wala fast, single-page dashboard. Koi backend/database nahi — sab kuch browser me process hota hai, isliye filtering aur rendering fraction of seconds me hoti hai.

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
- BO-wise cumulative performance table with totals
- Cumulative daily trend chart
- Flexible date filter:
  - **Upto date**: start date se kisi bhi selected date tak
  - **Date range**: kisi bhi from–to range ka data
  - **Single date**: sirf ek din ka data
- BO-wise filter (All BOs ya koi ek BO)
- **Export Excel**: currently filtered data `.xlsx` me download
- **Export PDF**: browser print dialog se PDF save (print-optimized layout)

## Local run

```bash
# koi bhi static server
python -m http.server 8000
# phir browser me: http://localhost:8000
```

Ya simply `index.html` ko browser me open karein.

## Render par host

1. GitHub repo ko Render se connect karein.
2. New → **Static Site** select karein.
3. Build command: **kuch nahi** (blank)
4. Publish directory: `.` (root)
5. Deploy — site live ho jayega aur har deploy par CDN se libraries load hongi.

Performance ke liye dashboard fully client-side hai: Excel parsing ek baar hota hai, uske baad date/BO filters turant (milliseconds me) re-render karte hain.
