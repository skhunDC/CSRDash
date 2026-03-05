# PLT CSR Dashboard

PLT CSR Dashboard is a Google Apps Script + HTMLService web app designed for TV display in store operations spaces. It visualizes weekly sales comparisons, CSR productivity, schedule/off-days, recognition, competitions, and end-of-day cleaning completion.

## Deploy as a Web App

1. Open the Apps Script project and paste files from root (`Code.gs`, `index.html`, etc.).
2. Save and run `ensureDatabase` once (to initialize Sheets schema/seed data).
3. Deploy:
   - **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone in your domain** (or appropriate workspace setting)
4. Copy deployment URL.

## Database Creation & Location

- Database uses a dedicated Google Spreadsheet created automatically by `ensureDatabase()`.
- The Spreadsheet ID is stored in **Script Properties** key: `CSR_DASHBOARD_DB_ID`.
- If the property is missing, a new spreadsheet is created.

## Tabs and Schema

- `Sales_Weekly`: `WeekStart, Store, Year, Sales`
- `CSR_Performance`: `Date, Store, CSRName, Sales, Hours`
- `CSR_Schedule`: `Date, Store, CSRName, ShiftStatus`
- `CSR_Recognition`: `WeekStart, CSRName, Store, Quote`
- `CSR_Competitions`: `WeekStart, Store, Metric, CSRName, Value, UpdatedAt`
- `Cleaning_Checklist`: `Date, Store, Task, Completed, CompletedBy, CompletedAt`
- `Auth_Log`: `Timestamp, Email, Authorized`

Sample seed rows are inserted automatically when tabs are empty.

## TV Mode

- Default Dashboard tab renders a one-screen card grid for all KPIs.
- Store filter is available directly inside the weekly schedule card.
- Dark theme + larger typography supports ceiling-mounted TV readability.
- Data refreshes every 5 minutes using server-driven config (`getClientConfig`).

## Change Allowed Users

In `Code.gs`, update `CONFIG.ALLOWED_EMAILS`.

## Change Competition Metrics

Metrics are data-driven through `CSR_Competitions.Metric` values.
- Add new metric names directly from the UI metric input (or by adding rows in Sheets).
- Existing metrics appear as suggestions via datalist.

## Print Checklist / Standings

Open web app URL with query parameter: `?page=print`.
- Choose **Daily Checklist** or **Competitions Standings**.
- Use browser print.

## Local Tests

```bash
cd test
npm install
npm test
```

Tests validate utility helpers and payload transformation logic with pure Node execution.
