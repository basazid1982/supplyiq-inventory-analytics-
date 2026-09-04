# SupplyIQ — Supply Chain Inventory Analytics Platform

A full-stack inventory management system with a built-in supply chain analytics layer — inventory turnover, ABC/Pareto classification, reorder-point suggestions, and demand forecasting — built on Google Apps Script, Google Sheets, and vanilla JS.

**Live demo:** [add your deployed web app URL here]
**Demo login:** `admin@example.com` / `admin123`

---

## The problem

Most small operations run inventory out of spreadsheets or gut feel: nobody notices a SKU is trending toward a stockout until a customer asks for something that isn't there, and nobody can say *which* products are actually worth the shelf space they take up. SupplyIQ is a working demonstration of the layer that's usually missing — not just "what do we have," but "what should we do about it."

## What it does

**Operations (the CRUD engine)**
- Products, categories, suppliers, and customers with full CRUD and validation (no duplicate SKUs/barcodes, no negative pricing, stock bounds enforced)
- Purchases and sales with multi-line items that automatically adjust stock and write to an immutable inventory movement log
- Deleting a transaction reverses its stock impact rather than leaving the ledger inconsistent
- Barcode/SKU quick-add on the point-of-sale screen
- Role-based login, CSV export per table, JSON backup/restore

**Analytics (the differentiator)**
- **Inventory turnover ratio** — trailing 12-month COGS ÷ current inventory value at cost
- **ABC/Pareto classification** — every SKU ranked by revenue contribution and bucketed into A (top 80% of revenue), B (next 15%), C (last 5%) — the standard framework for deciding where to focus purchasing and cycle-count effort
- **Reorder-point suggestions** — daily sales velocity (trailing 90 days) × configurable supplier lead time + safety stock, flagging exactly which SKUs need a purchase order now
- **Demand forecasting** — a least-squares linear regression over trailing monthly revenue, projected one month forward

Every analytics card ships with a one-line methodology note in the UI — the formula is visible, not a black box.

## Why these specific metrics

These aren't arbitrary dashboard filler — they're the four questions every inventory-heavy business actually needs answered on a recurring basis: *how efficiently is capital tied up in stock moving (turnover)? Which products deserve the most attention (ABC)? What do I need to reorder before I run out (reorder point)? Where is demand headed (forecast)?* Building the tooling to answer them, not just the CRUD to store the data, was the point of this project.

## Tech stack

- **Backend:** Google Apps Script (JavaScript), Google Sheets as the data store
- **Frontend:** Vanilla HTML/CSS/JS, single-file SPA, no framework dependency
- **Architecture:** Layered backend (config → utilities → auth → data access → business logic → analytics → web entry point), every function returns a consistent `{ success, message, data }` shape
- **Known trade-off:** Sheets-as-database is deliberately lightweight for a fast, free, zero-infrastructure build — it's not the choice I'd make for a production multi-tenant SaaS (that would move to Postgres + a proper backend), and I can speak to that trade-off directly.

## Setup

1. Create a new Google Sheet → Extensions → Apps Script
2. Create `Code.gs` and paste in the backend code
3. Create an HTML file named exactly `Index` and paste in the frontend code
4. Run `setupSheets()` then `seedSampleData()` from the Apps Script editor once (approve permissions when prompted)
5. Deploy → New deployment → Web app → Execute as "Me", access "Anyone"
6. Log in with the demo admin credentials above

## What I'd change for a production version

- Move the database to Postgres with a real backend (Node/Next.js) for concurrency and scale
- Add proper multi-tenancy so this could serve many businesses, not one spreadsheet per customer
- Replace the naive linear-regression forecast with a seasonality-aware model
- Add historical inventory-value snapshots so turnover ratio uses a true average rather than a current-value proxy

---

Built by [Your Name] — [LinkedIn URL] — [contact]
