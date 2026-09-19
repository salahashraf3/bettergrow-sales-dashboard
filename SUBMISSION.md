## Submission note

I built an upload-first sales dashboard that validates the workbook as a complete transaction before replacing the visible state. It starts with no dummy results. A rejected replacement keeps the prior valid dashboard, overlapping reads are latest-wins, and selecting the same file again works. Workbook content remains in browser memory.

The contract is deliberately explainable: one row within the first 20 must contain all four headers; only documented aliases match; duplicate mappings and incomplete records fail with original Excel row numbers; numeric/currency text parsing is strict; and duplicate normalized salesperson + month records are rejected. Negative budgets and actuals are outside scope because this assessment does not model returns or credit notes. Zero budget is always `N/A` / `No budget`, and equality is on target before rounding.

The UI uses semantic HTML, responsive CSS, vanilla JavaScript and SheetJS. A local IntersectionObserver system supplies AOS-style reveal motion with fallback and reduced-motion support. `node tests.js` covers each edge-case row in the agreed validation contract.

## 10-minute walkthrough

1. Upload-first flow (1 min): first-load empty state, no dummy dashboard, and valid upload.
2. Transactional state (1 min): reject a bad replacement while retaining valid data; reselect the same file; explain latest-wins race protection.
3. Header contract (2 min): first 20 rows, one complete header row, explicit aliases, duplicate mappings and Excel row numbers.
4. Data validation (2 min): empty/header-only sheets, strict amounts and AED text, incomplete rows, negatives and duplicate salesperson-month pairs.
5. Business logic (2 min): normalized salesperson grouping across distinct months, equality as on target and zero-budget `N/A` / `No budget`.
6. UI and tests (2 min): responsive KPI/chart/table, local reveal motion and reduced motion; run `node tests.js` and discuss future ERPNext/server controls.
