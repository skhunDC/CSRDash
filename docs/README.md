# Dublin Cleaners CSR Dashboard

## Overview
This repository contains a production-oriented Google Apps Script HTMLService web app for a mounted TV CSR dashboard. The app is designed for a 1920×1080 display, uses a branded dark presentation layer, refreshes itself every 60 seconds, and keeps all operational content in Google Sheets so the team can update data without redeploying the app.

## Design approach
The dashboard is intentionally split into four visual zones that can be read from a distance:

1. **Hero sales zone** for weekly sales, year-over-year delta, and store comparison.
2. **CSR performance zone** for yesterday’s ranked dollars-per-hour productivity.
3. **Task execution zone** as the primary operational focus.
4. **Support zone** for schedule exceptions, competitions, and employee recognition.

The visual system uses:
- a dark atmospheric background to reduce TV glare
- large display typography for fast readability
- high-contrast color accents for status recognition
- light animation for refreshes and task changes without becoming distracting
- Dublin Cleaners branding and logo placement in the main hero and unauthorized state

## Authentication and authorization
### Authentication flow
This app uses the most reliable built-in pattern available for Google Apps Script HTMLService web apps:
- deploy the web app as **Execute as: User accessing the web app**
- restrict deployment access appropriately in Apps Script
- use `Session.getActiveUser().getEmail()` server-side to identify the signed-in Google account
- authorize requests only if the email matches the approved allowlist

### Allowed users
- `skhun@dublincleaners.com`
- `skhun1@dublincleaners.com`
- `ss.sku@gmail.com`

### Authorization behavior
- The allowlist is enforced on the server in `Code.gs`
- Unauthorized users never receive dashboard data from server-side methods
- The initial HTML template receives only an authorization state object
- Unauthorized users see a branded Unauthorized screen instead of the app
- Task update actions and refresh actions also re-check authorization server-side

### Important deployment note
Because Google Apps Script identity behavior varies by Workspace/domain policy, the deployment should be configured and tested with the final target accounts. The manifest is set to `executeAs: USER_ACCESSING` and `access: ANYONE` so both the two Dublin Cleaners accounts and the approved Gmail account can reach the sign-in boundary, while the server-side allowlist remains the real access control.

## Google Sheets data model
The app stores data in a spreadsheet named **CSR Dashboard Data**. The spreadsheet is created automatically if it does not already exist. Its ID is persisted in Script Properties under `CSR_DASHBOARD_SPREADSHEET_ID`.

### Auto-created sheets
If missing, the app creates these sheets:
- `Sales`
- `CSR_Performance`
- `Tasks`
- `Schedule`
- `Competitions`
- `Employee_Of_Week`
- `Employees`

### Sheet schemas
#### Sales
| Column | Purpose |
| --- | --- |
| Week Label | Current reporting week label |
| Store | Store name |
| Current Year Sales | Current year weekly sales |
| Last Year Sales | Prior year weekly sales |

#### CSR_Performance
| Column | Purpose |
| --- | --- |
| Date | Source date |
| CSR | Employee name |
| Sales | Dollars sold |
| Hours | Hours worked |

#### Tasks
| Column | Purpose |
| --- | --- |
| Display Order | Render order on the dashboard |
| Task | Task name |
| Assigned CSR | Owner |
| Completed | TRUE/FALSE toggle |
| Updated At | Timestamp of latest dashboard toggle |

#### Schedule
| Column | Purpose |
| --- | --- |
| Date | Scheduled date |
| CSR | Employee name |
| Status | `OFF` or working status |
| Notes | Optional note |

#### Competitions
| Column | Purpose |
| --- | --- |
| Category | `Conversions`, `Patio Signups`, or `Alterations` |
| CSR | Employee name |
| Value | Current score |
| Goal | Goal for progress bar |
| Notes | Optional coaching note |

#### Employee_Of_Week
| Column | Purpose |
| --- | --- |
| Name | Employee name |
| Quote | Optional fallback quote used only if the daily quote API is unavailable |
| Image URL | Optional image |
| Highlight | Supporting recognition copy |

#### Employees
| Column | Purpose |
| --- | --- |
| CSR | Active CSR roster name |

## Application files
- `Code.gs`: server-side Apps Script logic, authorization, sheet setup, data shaping, and cached daily quote API integration
- `index.html`: main TV dashboard container
- `styles.html`: shared visual system and layout styles
- `scripts.html`: client-side rendering, refresh, and interactions
- `print.html`: printable simplified summary view
- `appsscript.json`: manifest, runtime, web app mode, and scopes

## Setup and deployment
1. Create or open the Apps Script project.
2. Copy these files into the project root.
3. Deploy as a web app.
4. Use **Execute as: User accessing the web app**.
5. Restrict access to the intended user group in the deployment settings.
6. Open the deployed app once as an authorized user to let the spreadsheet auto-create.
7. Populate the generated spreadsheet with live operational data before using the dashboard in production.

## Usage notes
- The dashboard refreshes every 60 seconds automatically.
- Task owners are edited in Google Sheets.
- Task completion can be toggled on the dashboard and is persisted back to Sheets.
- The `Employees` sheet is prefilled with the active CSR roster for Regina, Shelly, Nellie, Demetria, Dipali, Heather, Lynn, Lisa, Angela, Karmen, Omar, Kaylee, Ingrid, Kelly, Brandy, and Cheyenne.
- The printable summary view is exposed from the same web app deployment.

## Testing approach
The `test` directory contains a lightweight Node-based validation suite focused on static and structural checks for this repository.

It verifies:
- required project files exist
- the Apps Script manifest contains the expected web app settings and scopes
- `Code.gs` includes required allowlisted users, sheet names, and task labels
- HTML partials include key rendering hooks for authentication, auto-refresh, and task interactions

### Run tests
```bash
cd test
npm install
npm test
```

## Assumptions and implementation decisions
- The app uses Google Sheets as the only persistence layer to keep editing simple for operations staff.
- Sales, schedule, competition, performance, and employee-recognition sheets are created with headers only so live data can be entered directly.
- The web app is optimized for TV display first, not for editing-intensive mobile workflows.
- The print view is intentionally simplified for managers who need a quick paper-friendly task summary.
- The app uses progressive enhancement in the browser, but all sensitive authorization decisions remain on the server.

## Daily service quote behavior
- The Employee of the Week panel now requests one quote per day from a free quote API on the server side.
- The app tries They Said So management quote of the day first, then falls back to ZenQuotes if needed.
- The response is cached in Apps Script properties/cache for the current date so the TV refresh cycle does not repeatedly call the external API.
- If both APIs are unavailable, the `Employee_Of_Week` sheet `Quote` cell is used as the on-screen fallback.
