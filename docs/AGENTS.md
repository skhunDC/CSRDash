# Development Notes

## Project intent
This repository is for a polished Google Apps Script HTMLService CSR dashboard used on a mounted TV at Dublin Cleaners. Favor clarity, maintainability, and direct business usefulness over framework-heavy abstractions.

## Implementation guidelines
- Keep authentication checks server-side in Apps Script methods.
- Use Google Sheets as the editable data source unless a future requirement explicitly changes persistence.
- Preserve the TV-first visual hierarchy: hero sales, CSR performance, task focus, then support modules.
- Prefer readable Apps Script utilities and small pure formatting helpers over complex layers.
- Maintain branded, intentional UI treatment for both authorized and unauthorized states.

## Frontend guidance
- Optimize first for 1920×1080, large typography, and no-scroll presentation.
- Use motion sparingly and purposefully for refreshes, status changes, and hierarchy.
- Keep the Today’s Tasks section prominent because it is the operational priority.
- When adding UI, prefer composition and hierarchy over adding more boxed widgets.

## Testing guidance
- Keep tests friction-light and runnable with plain `npm test` from the `test` directory.
- Favor structural and regression-oriented checks that help protect Apps Script-specific files.
- Add or update tests whenever auth rules, sheet schema, or core UI structure changes.

## Documentation guidance
- Document how the auth model works in Apps Script terms, including deployment assumptions.
- Document any new sheet schema or operational workflow changes.
- Leave room for future enhancements; do not impose unnecessary process constraints.
