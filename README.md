# Bettergrow Sales Performance Dashboard

[Live demo](https://salahashraf3.github.io/bettergrow-sales-dashboard/)

A lightweight web app that accepts an Excel workbook, validates and aggregates its first worksheet in the browser, and shows headline metrics, a target-status table, and a budget-vs-actual chart.

## Run locally

Serve this folder with any static server, for example `npx serve .`, then open the local URL. The dashboard starts with realistic sample data. Upload `sample-sales-data.xlsx` to test the complete flow.

## Expected Excel columns

`Salesperson`, `Month`, `Budget Amount`, `Actual Sales`. Column names are matched case-insensitively and common spacing differences are accepted. Missing columns, empty workbooks, invalid or negative amounts, unsupported files, and files over 10 MB produce clear validation messages.

## Technical approach

- Semantic HTML and responsive CSS, without a framework or build pipeline.
- Vanilla JavaScript for validation, aggregation, state, table rendering, and the chart.
- SheetJS bundled locally for `.xlsx` and `.xls` parsing.
- In-memory processing only; the workbook never leaves the browser.

Rows are grouped by salesperson, so one person can appear across many months. Budget and actual amounts are summed before calculating achievement and status. Actual greater than or equal to Budget is above target. Zero-budget rows cannot cause a divide-by-zero failure.

## With more time

Add month/team filters, exports, tests, saved history with user roles, and a secure ERPNext API integration. A production version would also need server-side validation, audit logs, configurable currency, and configurable target rules.
