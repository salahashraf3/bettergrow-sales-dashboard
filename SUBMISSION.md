## Submission note

I built a lightweight sales-performance dashboard that validates an uploaded Excel workbook, aggregates monthly rows by salesperson, and turns the result into headline KPIs, a target-status table, and a budget-versus-actual chart. I chose semantic HTML, responsive CSS, vanilla JavaScript, and a locally bundled SheetJS parser because the brief does not need a backend, which keeps the solution fast, private, easy to deploy, and simple to explain. I used AI to break the brief into validation, aggregation, visualisation, and test cases, then reviewed the logic and adjusted the data model, error handling, accessibility, and responsive design myself. With more time, I would add month/team filters, automated tests, exports, user roles, upload history, and a secure ERPNext API integration.

## 10-minute walkthrough

1. Problem and user flow (1 min): built-in sample, then upload the standard sales workbook.
2. Architecture (1 min): static app, browser-only state, local SheetJS parser, separate HTML/CSS/JS.
3. Workbook parsing (2 min): first worksheet, normalized column matching, validation, file guards.
4. Business logic (2 min): group by salesperson, total Budget and Actual, Achievement %, Actual ≥ Budget target rule, zero-budget handling.
5. UI and chart (2 min): KPIs, accessible bar chart, badges, AED formatting, responsive table, privacy note.
6. Test and next steps (2 min): upload sample workbook, demonstrate an invalid workbook, show mobile layout, discuss filters, tests, ERPNext API, roles, history, and export.
