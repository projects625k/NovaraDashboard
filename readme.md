# Novara Business Intelligence

Internal sales & management dashboard for **Novara Properties**. One static `index.html`, one `data.json`, no backend required — deployable on GitHub Pages in minutes.

```
index.html   the complete application (UI, calculations, Excel parser, employee module, GitHub publisher)
data.json    sales transactions + data-quality report, generated from the Excel workbook
convert.py   command-line converter: Excel -> data.json (same rules as the in-browser upload)
readme.md    this file
```

## 1. Quick start (local)

Open `index.html` through a local web server (browsers block `fetch('data.json')` from `file://`):

```bash
python -m http.server 8000     # then open http://localhost:8000
```

If no `data.json` is found the dashboard shows a welcome screen — upload the Excel file from there.
Without a server you can still open `index.html` directly and use **Data → Upload Excel**; the data is cached in the browser (IndexedDB) and reloads next time.

## 2. Updating the numbers

There are three ways. All produce the same `data.json`.

| Method | When to use |
|---|---|
| **Data tab → Upload Excel** | Daily use. Replaces sales data in the browser, keeps employees/settings. |
| **Data tab → Connect GitHub (auto-publish)** | Live site. Every upload writes `data.json` into your repo; GitHub Pages redeploys automatically. |
| **`python convert.py CLOSURE_DASHBOARD_new.xlsx`** | Scripted / offline. Writes `data.json` next to the script; commit it. Requires `pip install openpyxl`. |

### Excel requirements
- Data on the **first sheet**, headers on row 2 (row 1 is a title). Header **names** are what matter, not column letters — columns can move.
- Required headers: `AGENT NAME`, `DATE OF CLOSURE`, `DEVELOPER`, `Net Revenue for Novara`, `TEAM`.
- Optional, used automatically when present: Property, Unit No, Client Name, RM, Source, Mode of Payment, Token Paid, DP/Doc Status, Property Value, Commission %, Commission Payout, Passback, VAT, Gross Commission, % to Agent / Team Leader / Sales Head / Telecaller, Invoice Number, Invoice Date, Year, Month.
- A sheet named like `2025` with two columns (Category | Amount) is imported as an expense summary (marketing ROAS panel on the Data tab).
- **Download sample Excel** on the Data tab produces a template with the exact layout.

### Cleaning rules (applied identically in `convert.py` and the browser)
- Blank rows and junk rows (no agent, date or revenue) are skipped.
- Revenue strings such as `AED 25,000` or `(1,200)` are converted to numbers; blank/invalid revenue → 0 and counted in Data Quality.
- Dates: Excel serials, real dates, `DD/MM/YYYY`, `MM/DD/YYYY`, `YYYY-MM-DD`, `11 JUN 2025`. Rows with no date but a `YEAR` + `MONTH` get an estimated first-of-month date (flagged `est.`).
- Agent names are trimmed/upper-cased and merged via aliases (Settings → Agent name aliases), e.g. `BAZEED KHAN → BAZEED`, `SHOAIB → SHUAIB`.
- Missing Developer/Team → `UNASSIGNED` (visible, never silently dropped).
- Possible duplicates (same agent + date + developer + unit + client + revenue) are kept but flagged.
- Column `Total Revenue For Novara` is a running total in the workbook and is intentionally ignored.

## 3. Deploy on GitHub Pages + auto-publish

1. Create a repository (e.g. `novara-bi`) and commit `index.html`, `data.json`, `convert.py`, `readme.md`.
2. Settings → Pages → Source: *Deploy from a branch* → `main` / root. Your dashboard is live at `https://<owner>.github.io/novara-bi/`.
3. Create a **fine-grained personal access token**: GitHub → Settings → Developer settings → Personal access tokens → Fine-grained. Repository access: *only this repo*. Permissions: **Contents: Read and write**.
4. Open the live dashboard → **Data → Connect GitHub**: enter owner, repo, branch (`main`), path (`data.json`), paste the token, tick **Auto-publish after every Excel upload**, then **Test connection** and **Save**.
5. From now on: upload the Excel in the Data tab → the dashboard updates instantly *and* `data.json` is committed to the repo → GitHub Pages rebuilds (≈1 minute) → the live site shows the new figures. No manual download/upload.

Optional: set *Also publish backup to* (e.g. `backup/novara_backup.json`) to commit employee master data alongside.

**Security note.** The token lives only in the browser's IndexedDB and is sent only to `api.github.com`. Anyone with access to that browser profile could publish to the repo; use a token limited to one repository and revoke it from GitHub when no longer needed. Untick *Remember token* to be prompted each session.

## 4. What persists where

| Data | Storage | Survives Excel re-upload | In backup |
|---|---|---|---|
| Sales transactions | `data.json` (or IndexedDB cache after upload) | replaced | no (regenerate from Excel) |
| Employees, photos, salaries, benefits, statuses, org relationships | IndexedDB `employees` | yes | yes |
| Agent-name → employee mappings | IndexedDB `mappings` | yes | yes |
| Targets | IndexedDB `targets` | yes | yes |
| Settings, thresholds, aliases, logo, theme | IndexedDB `settings` | yes | yes |
| GitHub config | IndexedDB `github` | yes | yes (token excluded) |

**Data → Download backup / Restore backup** exports and imports everything except sales data as JSON. Employee data is per-browser; move it between machines with the backup file, or migrate the stores to a database (each IndexedDB store maps 1:1 to a table — see `DB` in `index.html`).

## 5. Pages

- **Executive Summary** — KPI cards with period comparison (month → previous month, quarter → previous quarter, year → previous year, custom range → equal preceding range; no time filter → trailing 12 months vs prior 12), revenue trend (day/week/month/quarter/year, line/area/column, click to drill), cumulative revenue, revenue vs transactions, top agents/developers, team chart + donut, revenue by quarter and lead source, agent/team heatmap, management insights, top alerts, target achievement, leaderboard.
- **Teams** — team profile, trend, agent ranking, developer mix, distribution, heatmap, performance table with status, team leaderboard with growth.
- **Developers** — developer profile, trends, top agents/teams, projects, deal-size bands, commission waterfall (gross → splits → net), 2–5 developer comparison, leaderboard.
- **Agents** — profile card (photo upload, joining date, contract type, employment status), recency indicator with configurable thresholds, 0–100 performance score (revenue rank 30 %, transaction rank 20 %, recency 20 %, consistency 15 %, growth 15 %; new agents are not penalised), inactivity alerts, lifetime/period KPIs, charts, searchable/sortable/paginated sales history with Excel/CSV export, link-to-employee mapping.
- **Leaderboards** — agents, teams, developers with growth; export.
- **Alert Centre** — critical inactivity, revenue decline (last complete month and quarter), concentration risk, developer opportunities, team records. Agents flagged resigned/terminated or inactive for over a year are excluded from inactivity alerts.
- **Employees** — add/edit/delete (with confirmation), search/filter, table/card view, summary metrics, export. Saving an Agent profile auto-creates a linked employee.
- **Organisation Chart** — built from *Reporting manager*; zoom/pan, expand/collapse, fullscreen, search; click to open profile.
- **Data** — Excel upload, data.json download, sample Excel, GitHub auto-publish, data-quality panel, column mapping, backup/restore, expense/ROAS panel.
- **Settings** — company name, logo, currency, date format, default year, theme, activity thresholds, agent aliases, monthly targets (company / team / agent).

Global filters (Year, Quarter, Month, date range, Agent, Team, Developer — all multi-select, with removable chips and one-click reset) and global search apply across pages.

## 6. Verification performed

- Total Net Revenue ties to the workbook sum of column AB (AED 22,674,144.44 over 658 closures, 28 Jul 2023 → 07 Aug 2026).
- `convert.py` and the browser parser produce identical transaction sets (field-by-field diff = 0).
- Year / quarter / month / date-range / agent / team / developer filters and reset verified in a headless browser; no console errors; no `NaN`/`undefined` rendered; layout holds at 1440 px and 900 px.
- Employee records survive page reload and Excel re-upload; org chart rebuilds from reporting manager.

## 7. Stack

Vanilla JS (ES2020), [Apache ECharts 5](https://echarts.apache.org) for charts, [SheetJS](https://sheetjs.com) for Excel read/write, IndexedDB for persistence, GitHub Contents API for publishing. Fonts: Inter / Inter Tight / JetBrains Mono. Both libraries load from cdnjs; to run fully offline, download them and change the two `<script src>` lines at the top of `index.html`.
