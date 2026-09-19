# Bettergrow Sales Performance Dashboard

[Live demo](https://salahashraf3.github.io/bettergrow-sales-dashboard/)

A browser-based business dashboard that accepts an Excel workbook, validates the first worksheet as one complete transaction, and then shows KPIs, a target-status table and a budget-vs-actual chart. The first visit contains no dummy metrics, rows or chart.

## Run and test locally

Serve this folder with any static server, for example `npx serve .`, then open the local URL. Run the focused validation suite with `node tests.js`.

## Workbook contract

All four required headers must appear together in one row within the first 20 Excel rows. Matching is case-insensitive and uses this explicit alias map only:

- `Salesperson`: Salesperson, Sales Person, Salesperson Name, Sales Rep, Sales Representative
- `Month`: Month, Sales Month, Period
- `Budget Amount`: Budget Amount, Budget, Target, Target Amount
- `Actual Sales`: Actual Sales, Actual, Actual Sales Amount, Sales Amount

Fuzzy matching is intentionally not used. Two columns in the selected header row cannot map to the same field. Blank rows are ignored. Every nonblank row below the header must be complete, and errors retain its original Excel row number.

## Validation and business rules

- `.xlsx` and `.xls` extensions are checked case-insensitively. Empty files and files over 10 MB are rejected before parsing, but every accepted extension is still parsed as workbook content.
- Unreadable, damaged or password-protected input gets a cautious read error rather than claiming a cause the parser cannot prove.
- Empty sheets, missing required columns and header-only sheets have distinct messages.
- Blank amounts are missing, not zero. Numeric Excel cells are valid. Text is accepted only as a strict plain number or AED amount such as `AED 1,000.50`. Booleans, nonfinite values, partial strings, ambiguous separators, symbols and other/mixed currencies are rejected.
- Negative Budget and Actual Sales are rejected. Returns and credit notes are outside this assessment's scope.
- Salesperson and month keys are trimmed, repeated spaces are collapsed and comparisons are case-insensitive. The cleaned first display name is preserved.
- A repeated normalized salesperson + month pair is an error. The same salesperson in different months is valid and aggregates into the salesperson summary.
- Actual equal to Budget is `On target`; amounts are compared before percentage rounding.
- Any zero budget is shown as achievement `N/A` with status `No budget`, including zero budget with positive actual. It is not counted on or below target.
- Upload validation is transactional: an invalid replacement leaves the prior valid dashboard visible. Overlapping asynchronous uploads are latest-wins, stale reads cannot overwrite newer data, and the same file can be selected again.

## Technical approach

Semantic HTML and responsive CSS; vanilla JavaScript for parsing state, validation, aggregation, table and chart rendering; SheetJS for local `.xlsx`/`.xls` parsing. Workbook contents stay in browser memory. A self-contained IntersectionObserver reveal system provides AOS-style motion, has an older-browser fallback and respects `prefers-reduced-motion`.

## With more time

Add month/team filters, exports, saved history with user roles and a secure ERPNext API integration. A production version would also need server-side validation, audit logs and configurable currency/target policies.
