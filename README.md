# ALIMCO Export Intelligence Dashboard

AI-powered export pipeline dashboard with live Claude analysis, Leaflet world map, urgent alerts, and interactive charts.

---

## Project Structure

```
dashboard/
├── app.py                        ← Flask backend (all endpoints + AI)
├── dashboard.html                ← Single-file frontend (served by Flask)
├── "exports dashboard.xlsx"      ← Your Excel data file (rename if needed)
├── .env                          ← Your API key (create this yourself)
├── requirements.txt
├── Procfile                      ← For Railway / Render deployment
├── runtime.txt
└── README.md
```

---

## Setup (Local)

### 1. Rename your Excel file
Your uploaded file may have a typo. Rename it exactly to:
```
exports dashboard.xlsx
```
Place it in the same folder as `app.py`.

> The app also auto-detects `exports_dashbaord.xlsx` and `exports_dashboard.xlsx` as fallbacks.

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Create your `.env` file
```
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxxxxxxxx
```
Get your key from: https://console.anthropic.com

### 4. Run the backend
```bash
python app.py
```

### 5. Open the dashboard
Open your browser and go to:
```
http://localhost:5000
```

> **Important:** Do NOT open `dashboard.html` directly as a file.
> It must be served by Flask at `http://localhost:5000` so the API calls work.

---

## Excel File Requirements

- **Filename:** `exports dashboard.xlsx`
- **Sheet name:** `Leads Data` (exact, case-sensitive)
- **Header row:** Row 3 (rows 1–2 are the title banner — keep them as-is)

Required columns (exact names):

| Column | Type |
|--------|------|
| SNO. | Integer |
| Company Name | Text |
| Contact Person | Text |
| Last Contacted | Date |
| Order # | Text |
| Order Type | Text |
| Order Value (₹ INR) | Number |
| Order Value ($ USD) | Number |
| Lead Source | Text |
| Lead Status | Text |
| Deal Stage | Text |
| Email Address | Text |
| Country | Text |
| City | Text |
| Pin Code | Text |
| Contact Number | Text |
| Notes | Text |
| Last Updated | Date |

---

## Features

- **AI Executive Briefing** — Claude analyses your pipeline and generates a 6-section strategy report
- **Urgent Alert Popups** — Claude scans Notes column and surfaces CRITICAL/HIGH/MEDIUM alerts
- **Leaflet World Map** — Dark-tile map with circle markers sized by pipeline value
- **5 Interactive Charts** — Country pipeline bar, lead status doughnut, pipeline trend line, deal stage doughnut, lead source horizontal bar
- **Sortable Table** — All 18 columns, click any header to sort
- **Filter Bar** — Filter by Country, Status, Stage, Type, Source — all charts and map update live
- **Refresh Button** — Re-reads Excel and re-runs AI without restarting the server

---

## API Endpoints

| Method | Route | Description |
|--------|-------|-------------|
| GET | `/` | Serves dashboard.html |
| GET | `/api/leads` | All leads + AI insights + alerts |
| GET | `/api/alerts` | Just the alerts array |
| GET | `/api/map` | Country aggregates with lat/lng |
| POST | `/api/refresh` | Reloads Excel + re-runs AI |
| GET | `/api/health` | Status check |

---

## Deploy to Railway (Share Online)

1. Push this folder to a GitHub repository
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Select your repo
4. In Railway settings → Variables, add:
   ```
   ANTHROPIC_API_KEY = sk-ant-...
   ```
5. Upload `exports dashboard.xlsx` via Railway's file manager
6. Railway reads `Procfile` and deploys automatically
7. Your public URL: `https://yourapp.up.railway.app`

Share that URL — anyone opens it in a browser, no setup needed.

### Alternative: Render.com
Same steps as Railway. Free tier sleeps after 15 min idle (first load after sleep is slow). Paid tier stays always-on.

---

## Fallback Behaviour

| Situation | What happens |
|-----------|-------------|
| Excel file missing | 10-row demo data loads automatically |
| ANTHROPIC_API_KEY not set | Dashboard works without AI; briefing shows placeholder |
| Backend unreachable | Frontend shows demo data with an error banner |

---

## Troubleshooting

**Charts not showing** — Make sure you open `http://localhost:5000`, not `dashboard.html` directly.

**AI says "unavailable"** — Check `.env` file exists and `ANTHROPIC_API_KEY` is set correctly.

**Excel not loading** — Confirm filename is exactly `exports dashboard.xlsx` and sheet is `Leads Data`.

**Map markers missing** — Countries must match the spelling in the app's coordinate dictionary. Common issue: `tanzania` (lowercase) — the app normalises case automatically.
