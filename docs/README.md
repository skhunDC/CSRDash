# PLT CSR Dashboard

PLT CSR Dashboard is a Google Apps Script + HTMLService web app for TV-friendly CSR operations tracking with in-app admin editing.

## Deploy as a Web App

1. Open the Apps Script project and paste files from root (`Code.gs`, `index.html`, etc.).
2. Save and run `ensureDatabase` once (initializes all sheet tabs + seed rows).
3. Deploy Web App and share in your workspace.

## Data Storage (Google Sheets Tabs)

Core tabs:
- `Sales_Weekly`
- `CSR_Performance`
- `CSR_Schedule`
- `CSR_Recognition`
- `Cleaning_Checklist`
- `Auth_Log`

Competition system (dynamic, admin-editable):
- `Competition_Categories`: `CategoryKey, CategoryName, Enabled, SortOrder, Goal, Notes, UpdatedAt`
- `Competition_Entries`: `WeekStart, Store, CSR, CategoryKey, Value, Notes, UpdatedAt`

Live links system:
- `App_Links`: `Key, Label, Url, UpdatedAt`
- seeded key: `PATIO_CUSHION_SIGNUP`

## Admin Workflows (No Redeploy Needed)

### Add / rename / remove competition tabs
1. Open **Metrics Admin**.
2. Add or edit rows in the category table.
3. Set `Enabled` to show/hide tabs.
4. Save.

Enabled categories immediately render as top navigation tabs and drive each competition table.

### Edit Patio Cushion Signup link
1. Open **Links Admin**.
2. Update `Label` and `URL` for `PATIO_CUSHION_SIGNUP`.
3. Save.

Dashboard + Patio Signups tab buttons update immediately without code or redeploy changes.

## Runtime Behavior

- Category tabs default to TV read-only view.
- Use **Manage** mode inside a category tab to inline edit `Value` and `Notes`.
- Save writes in batch via `saveCompetitionEntries(payload)`.
