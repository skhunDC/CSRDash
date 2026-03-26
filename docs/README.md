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
- Dublin Cleaners branding and logo placement in the main hero and fallback error states

## Authentication and authorization
### Dashboard access flow
The dashboard itself is now intentionally public:
- deploy the web app as **Execute as: Me / USER_DEPLOYING**
- publish it with `access: ANYONE`
- let server-side Apps Script methods read and update the dashboard spreadsheet on behalf of the deployment owner
- do not require a Google sign-in just to load the TV dashboard or refresh its data

### Source sheet access behavior
The **Open source sheet** link still uses Google's normal spreadsheet access controls:
- clicking it sends the browser to the Google Sheets URL for the backing file
- Google will prompt for sign-in if the visitor is not already authenticated
- the sheet only opens for accounts the file owner has explicitly granted access to
- users without permission will see Google's standard access-denied/request-access flow

### Important deployment note
Because the dashboard is public, anyone with the web app URL can view the board and use dashboard actions. The manifest is set to `executeAs: USER_DEPLOYING` and `access: ANYONE` so the mounted TV can load without authentication, while the separate Google Sheets link still respects the spreadsheet's own sharing permissions.

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
- `Code.gs`: server-side Apps Script logic, public dashboard data access, sheet setup, data shaping, and cached daily quote API integration
- `index.html`: main TV dashboard container
- `styles.html`: shared visual system and layout styles
- `scripts.html`: client-side rendering, refresh, and interactions
- `print.html`: printable simplified summary view
- `appsscript.json`: manifest, runtime, web app mode, and scopes

## Setup and deployment
1. Create or open the Apps Script project.
2. Copy these files into the project root.
3. Deploy as a web app.
4. Use **Execute as: Me** so the public dashboard can load with the deployment owner's spreadsheet access.
5. Set deployment access to **Anyone** for the dashboard URL, and control spreadsheet editing/viewing separately through Google Sheets sharing.
6. Open the deployed app once after deployment to let the spreadsheet auto-create.
7. Populate the generated spreadsheet with live operational data before using the dashboard in production.

## Usage notes
- The dashboard refreshes every 60 seconds automatically.
- Task owners are edited in Google Sheets.
- Task completion can be toggled on the dashboard and is persisted back to Sheets.
- The `Employees` sheet is prefilled with the active CSR roster for Regina, Shelly, Nellie, Demetria, Dipali, Heather, Lynn, Lisa, Angela, Karmen, Omar, Kaylee, Ingrid, Kelly, Brandy, and Cheyenne.
- The printable summary view is exposed from the same web app deployment.
- The **Open source sheet** link intentionally hands off to Google Sheets, which will prompt for sign-in and enforce the file owner's sharing permissions.

## Testing approach
The `test` directory contains a lightweight Node-based validation suite focused on static and structural checks for this repository.

It verifies:
- required project files exist
- the Apps Script manifest contains the expected public web app settings and scopes
- `Code.gs` includes required public-access, sheet, and task configuration
- HTML partials include key rendering hooks for auto-refresh, sheet access messaging, and task interactions

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
- The print workflow is intentionally simplified into an in-dashboard modal so managers can open, review, and print a paper-friendly task summary without leaving the main screen.
- The app uses progressive enhancement in the browser, while the linked Google Sheet continues to rely on Google-managed authentication and file permissions.

## Daily service quote behavior
- The Employee of the Week panel now requests one quote per day from a free quote API on the server side.
- The app tries They Said So management quote of the day first, then falls back to ZenQuotes if needed.
- The response is cached in Apps Script properties/cache for the current date so the TV refresh cycle does not repeatedly call the external API.
- If both APIs are unavailable, the `Employee_Of_Week` sheet `Quote` cell is used as the on-screen fallback.

## Infinite list ticker in Google Apps Script HTMLService
Because this project runs in Google Apps Script HTMLService (no React build step), ticker behavior is implemented directly in `scripts.html` + `styles.html`:
- each overflow list viewport is wrapped with a `.ticker-track`
- the original list is duplicated with DOM cloning
- track animates `translateY(0)` to `translateY(-50%)` with linear timing
- hover pauses animation
- top and bottom fade overlays are applied with pseudo-elements

This keeps the behavior seamless and deployment-safe for the current GAS environment without introducing `.jsx`/bundling requirements.
