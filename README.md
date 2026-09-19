# Bettergrow Sales Performance Dashboard

[Live demo](https://salahashraf3.github.io/bettergrow-sales-dashboard/)

A browser-based business dashboard that accepts an Excel workbook, validates and aggregates its first worksheet, and shows headline metrics, a target-status table, and a budget-vs-actual chart.

## Run locally

Serve this folder with any static server, for example `npx serve .`, then open the local URL. The first visit is an intentional empty state: no sample or placeholder metrics are presented as real results. Upload a valid `.xlsx` or `.xls` workbook to reveal the dashboard.

## Expected Excel columns

`Salesperson`, `Month`, `Budget Amount`, `Actual Sales`. Header matching is case-insensitive, ignores punctuation and spacing, accepts a small set of clear aliases, and can locate the header within the first 20 rows. Empty rows are skipped.

Validation covers unsupported or empty files, the 10 MB size limit, unreadable/password-protected workbooks, empty sheets, missing columns, missing names/months, malformed numeric cells and negative amounts. Errors preserve the upload state so the user can recover with another file.

## Business rules

- Rows are grouped by salesperson case-insensitively, so repeated monthly rows and harmless name-case differences aggregate together.
- Budget and actual amounts are summed before achievement is calculated.
- Actual greater than or equal to Budget is on target.
- A zero budget with zero actual is `0.0%`; a zero budget with positive actual is shown as `N/A` rather than a misleading infinite percentage.
- Currency-formatted values and comma separators are accepted; ambiguous or partial numeric strings are rejected.

## Technical approach

- Semantic HTML and responsive CSS, without a framework or build pipeline.
- Vanilla JavaScript for validation, aggregation, state, table rendering and the chart.
- SheetJS for `.xlsx` and `.xls` parsing; the workbook contents stay in browser memory.
- A self-contained IntersectionObserver reveal system provides AOS-style motion, includes an older-browser fallback and honours `prefers-reduced-motion`.
- No dummy dataset is bundled or rendered.

## With more time

Add month/team filters, automated browser tests, exports, saved history with user roles, and a secure ERPNext API integration. A production version would also need server-side validation, audit logs, configurable currency and configurable target rules.
