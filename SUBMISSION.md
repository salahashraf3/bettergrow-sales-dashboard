## Submission note

I built a focused sales-performance dashboard that validates an uploaded Excel workbook, aggregates monthly rows by salesperson, and turns the result into headline KPIs, a target-status table, and a budget-versus-actual chart. The app begins in a clear upload state and does not show dummy results. I chose semantic HTML, responsive CSS, vanilla JavaScript and SheetJS because the brief does not need a backend; workbook content stays in browser memory. The motion system is implemented locally with IntersectionObserver rather than an animation dependency, with an older-browser fallback and reduced-motion support. I used AI to break the brief into validation, aggregation, visualisation and test cases, then reviewed and adjusted the rules, error handling, accessibility, responsive design and edge-case behaviour myself.

## 10-minute walkthrough

1. Problem and user flow (1 min): explain the upload-first empty state, then load a valid workbook.
2. Architecture (1 min): static app, browser-only state, SheetJS parser, separate HTML/CSS/JS.
3. Workbook parsing (2 min): first worksheet, header search and normalization, file guards and row-level validation.
4. Business logic (2 min): case-insensitive grouping, duplicate/month aggregation, Budget and Actual totals, Actual ≥ Budget target rule, and explicit zero-budget handling.
5. UI and motion (2 min): executive KPIs, accessible chart, status badges, AED formatting, responsive table, local reveal animation and reduced-motion support.
6. Tests and next steps (2 min): demonstrate a valid upload and recovery from an invalid one; mention tests for empty/malformed sheets, varied headers, bad/negative values, duplicate rows, equality, zero budgets, repeated uploads and file guards; then discuss filters, exports, automated tests, ERPNext API, roles and history.
