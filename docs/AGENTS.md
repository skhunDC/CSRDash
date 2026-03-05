# PLT CSR Dashboard Architecture Notes

## High-Level Architecture

- **Client:** `index.html` shell + `styles.html` + `scripts.html` (vanilla JS).
- **Server:** `Code.gs` Apps Script endpoints.
- **Database:** Google Spreadsheet managed through Script Properties + generated tabs.

## Data Flow

1. Browser loads `index.html`.
2. Client calls `getAuthStatus()`.
3. If authorized, client requests:
   - `getDashboardData()`
   - `getCompetitionsData()`
   - `getChecklist(date, store)`
4. User updates push through write endpoints:
   - `saveCompetitionEntry(payload)`
   - `setChecklistItem(payload)`
5. Server writes to Sheets, clears cache, and UI re-renders.

## Auth Flow Steps

1. `getCurrentUserEmail_()` checks `Session.getActiveUser().getEmail()` then falls back to `Session.getEffectiveUser().getEmail()`.
2. Email is checked against `CONFIG.ALLOWED_EMAILS`.
3. Every attempt is appended to `Auth_Log` via `logAuthAttempt`.
4. Unauthorized users receive full-page unauthorized state and no app UI rendering.

## Expansion Notes

- Add store-level ACL logic if future permissions vary by location.
- Add historical trend charts by expanding `Sales_Weekly` to include week number and prior weeks.
- Add role-based admin tab for metric/task management.
- If scale grows, move heavy computed payloads into periodic materialized sheets.
