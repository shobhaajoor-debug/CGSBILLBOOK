# CGS Global Enterprises Bill Book

A local billing and payment-tracking web application for CGS Global Enterprises. It uses a vanilla HTML/CSS/JavaScript frontend, a Node.js/Express API, JSON ledgers for application records, and an Excel workbook generated from those ledgers. A local PDF service extracts vendor-invoice fields; it does not call an external AI service.

This file is the repository map for developers and future coding agents. Read it before changing data flow, billing calculations, workbook synchronization, or extraction.

## At a glance

- **Main app:** `server.js`, normally at `http://localhost:3000`.
- **PDF extractor:** `backend_dataserve/server.js`, started by the main app, normally at `http://127.0.0.1:3001`.
- **Source payment tracker:** `CGS-Payment Tracker-Latest.xlsx` in the repository root.
- **Application data:** `data/bills.json` and `data/payments.json` are the record stores. `data/bills.xlsx` is generated from them.
- **Home page:** `public/dashboard.html`; the other pages are New Bill, Accounts, Bill History, and bill preview.
- **Tests:** `npm test` from the root runs the main test suite, including API, workbook, calculations, invoice PDF, and tracker import tests.

## Run locally

Requirements: Node.js 18+ and npm on Windows, macOS, or Linux.

```sh
npm install
npm install --prefix backend_dataserve
npm start
```

Open [http://localhost:3000](http://localhost:3000). On Windows, `start-app.bat` performs dependency checks and opens the app. Keep the server window open while using the app; press Ctrl+C there to stop it. Close `data/bills.xlsx` in Excel before bill/payment changes or workbook export.

The main server starts the PDF extractor as a child process. The extractor is local-only and accepts browser CORS requests from the main app on localhost. A typical browser UI operation uses the main server; do not start another server if port 3000 is already serving the app.

Useful environment variables:

| Variable | Default | Purpose |
| --- | --- | --- |
| `PORT` | `3000` | Main app HTTP port. |
| `DATA_DIRECTORY` | `<repo>/data` | Directory for the JSON ledgers and generated Excel workbook. |
| `PAYMENT_TRACKER_PATH` | `<repo>/CGS-Payment Tracker-Latest.xlsx` | Source tracker workbook. Set to `off` to disable tracker synchronization, or to a workbook path to use another file. |
| `EXTRACTOR_HOST` | `127.0.0.1` | PDF extractor bind address. |
| `EXTRACTOR_PORT` | `3001` | PDF extractor port. |

## User-facing workflows

### Overview

`public/dashboard.html` and `public/js/dashboard.js` show billed/received/outstanding totals, profit totals where known, monthly billing, and largest customer balances.

### Create a bill

`public/index.html` and `public/js/form.js` contain the New Bill form. The form captures seller and customer names and full addresses, invoice date, due days, vehicle number, line items, percentage increase, and payment status. It does **not** ask for purchase cost; new bills therefore store unknown cost and profit as `null`, not zero or the full invoice value.

Line amounts are calculated from quantity × excluding-tax rate. The total applies the configured percentage increase, deducts rows marked Discount, adds calculated GST, then rounds upward to a whole rupee. The server recalculates and validates submitted bill values; do not rely on client totals as authoritative.

Successful submission creates a sequential `CGS-####` bill ID, adds the bill and items to the JSON ledger and generated workbook, and stores an NIC-style E-Way Bill **draft** at `ewayBillDraft.payload`. The draft is not ready for government submission: required GST registration, normalized state/place/pincode, transport, unit, and tax classification details are not all collected. The application does not connect to or submit to the government E-Way Bill service.

### Extract a vendor PDF

The browser sends a selected PDF to `backend_dataserve/server.js` (`POST /api/extract`). `backend_dataserve/local-invoice-parser.js` parses the PDF text and best-effort extracts invoice values, items, seller, customer, and their full addresses. The fields are returned to the browser and also stored as structured drafts in `backend_dataserve/data/extracted-drafts.json`. The source PDF text is returned for display on the current page but is not persisted. Scanned/image-only PDFs need OCR and are not supported.

Extraction is best-effort. The user must review all populated fields, especially long addresses, GST/tax values, and totals, before submitting. Parser fixtures and tests are in `backend_dataserve/tests/extraction.test.js`. More extractor detail is in `backend_dataserve/README.md`.

### Accounts and receipts

`public/explore.html` and `public/js/explore.js` display aggregated customer balances and accept dated receipts. A receipt is applied to the customer's combined outstanding balance, not to a particular invoice; overpayments are rejected. Receipts are stored separately in `data/payments.json`.

### History, preview, and export

`public/history.html` and `public/js/history.js` list/search/filter bills, change a bill's Paid/Unpaid status, delete bills, and download an Excel export. A bill opened from history is rendered by `public/bill.html` and `public/js/bill.js`; the browser preview shows seller/customer addresses when present. **Download PDF** generates an invoice using `lib/invoice-pdf.js`. Imported tracker records have no original line items, so they show as tracker details and cannot produce a fabricated tax invoice.

## Data and Excel synchronization

### Which file is authoritative?

- `data/bills.json` is the application bill ledger, including bill details, line items, addresses, tracker source fields, and E-Way Bill drafts.
- `data/payments.json` is the dated customer receipt ledger.
- `data/bills.xlsx` is a generated workbook for viewing/export. The server regenerates its `Bills`, monthly, `Items`, `Payments`, and metadata sheets from the ledgers as needed. Do not treat manual edits to this generated workbook as bill changes.
- `CGS-Payment Tracker-Latest.xlsx` is a separate **source workbook** for imported tracker invoices.

Every main-app ledger read checks the source tracker's file signature. Once the tracker file has been saved in Excel, refreshing an app page makes changed source rows/values available through the app API. Tracker rows are replaced from the source workbook; app-created bills are preserved. Keep IDs unique: a tracker invoice conflicting with an app-created bill is reported as an error.

The importer is `scripts/import-payment-tracker.js`; it handles monthly sheets, varied headers, cached formula results, and duplicate invoice checks. It excludes non-invoice sheets and address columns. Tracker source rows do not provide original invoice line items or reliable receipt dates. Tracker values, including paid amounts and due terms, are imported for account/reporting use; imported addresses and invoice tax details must not be invented.

**Destructive import warning:** `node scripts/import-payment-tracker.js --replace` rebuilds `data/bills.json` from the tracker and resets `data/payments.json` to an empty array. Do not run it casually. Normal app startup/read synchronization does not clear dated payments or app-created bills. When changing the import or sync logic, review `lib/payment-tracker-sync.js` and its tests before touching production data.

New bills have no purchase cost unless a caller explicitly supplies one through the API. Consequently profit values can be null; UI/report changes must preserve that distinction. Imported tracker bills have historical purchase cost when the source workbook provides it.

## Repository map

| Path | Responsibility |
| --- | --- |
| `server.js` | Main Express app, bill/payment APIs, JSON ledgers, Excel workbook generation, tracker sync integration, PDF download, and startup of the extractor. |
| `public/dashboard.html`, `public/explore.html`, `public/history.html`, `public/index.html`, `public/bill.html` | Overview, customer accounts, bill history, bill form, and bill preview pages. |
| `public/js/dashboard.js`, `explore.js`, `history.js`, `form.js`, `bill.js` | Page-specific browser logic. |
| `public/css/style.css` | Shared styling and responsive layout. |
| `lib/payment-tracker-sync.js` | Refreshes tracker-sourced records while retaining app-created bills and receipts. |
| `scripts/import-payment-tracker.js` | Reads tracker workbooks and performs explicit initial/replacement imports. |
| `lib/tax-summary.js` | Calculates HSN-level taxable value and GST summary. |
| `lib/amount-in-words.js` | Formats invoice totals in Indian currency words. |
| `lib/invoice-pdf.js` | Builds downloadable tax invoice PDFs. |
| `lib/eway-bill.js` | Builds an incomplete, explicitly non-submittable E-Way Bill JSON draft from known bill data. |
| `backend_dataserve/server.js` | Local PDF upload/extraction API and structured draft persistence. |
| `backend_dataserve/local-invoice-parser.js` | Vendor-neutral invoice/date/party/address/item parsing. |
| `backend_dataserve/json-store.js` | Atomic JSON append helper used by extraction drafts. |
| `tests/billing.test.js` | Calculations, main API/workbook behavior, payment flows, invoice PDF, and E-Way Bill draft tests. |
| `tests/payment-tracker-import.test.js` | Source workbook import and edit-refresh tests. |
| `backend_dataserve/tests/extraction.test.js` | Parser, address, and structured-draft tests. |

## Main API routes

Routes are implemented in `server.js`; request validation and workbook/ledger writes are server-side.

| Method and path | Purpose |
| --- | --- |
| `GET /api/health` | Main app readiness/version. |
| `POST /api/bills` | Validate and save a new bill, items, and E-Way Bill draft. |
| `GET /api/bills` | List bills for history/dashboard. |
| `GET /api/bills/:billId` | Read bill detail for preview. |
| `GET /api/bills/:billId/pdf` | Download a generated invoice PDF (not available for tracker imports). |
| `DELETE /api/bills/:billId` | Delete a bill. |
| `PATCH /api/bills/:billId/status` | Set bill payment state to Paid or Unpaid. |
| `GET /api/bills/export` | Download the generated Excel workbook. |
| `GET /api/explore` | Customer balances, receipts, and account summary. |
| `POST /api/payments` | Record a dated receipt against a customer's balance. |

Extractor routes: `GET http://127.0.0.1:3001/api/health`, `GET /api/drafts`, and `POST /api/extract` with a multipart `pdf` field.

## Tests and change guidance

```sh
npm test
npm test --prefix backend_dataserve
```

The root suite uses temporary data directories for API/workbook integration tests. Optional PDF fixture tests are skipped when their local files are absent. Extractor tests can be run independently from `backend_dataserve`.

Before changing behavior:

1. Trace the full flow: page HTML → page JS → API route → JSON ledger → workbook/preview/PDF.
2. Keep derived totals and validation consistent between browser and server; the server is authoritative.
3. Preserve tracker versus app-created record distinctions and avoid fabricating tracker items/receipts.
4. Preserve full seller/customer addresses through extraction, UI, JSON storage, E-Way Bill draft, preview, and PDF. Do not persist source PDF text.
5. Add or adjust the smallest relevant test suite, then run `git diff --check`.
6. Do not overwrite or reset unrelated existing worktree changes.
