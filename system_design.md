# System Design: Financial Analysis & Profit Management System

An interactive, premium-grade financial intelligence dashboard designed for small-to-medium enterprises (SMEs) to track, analyze, and optimize company income, expenses, and profits.

---

## 1. System Overview

This system serves as a central hub for analyzing financial health. It translates raw income and expense transactions into actionable insights. By incorporating dynamic configurations for **currency** (defaulting to Japanese Yen `¥`) and **Value Added Tax (VAT)** (defaulting to `10%`), it provides localized, exact calculations for businesses operating globally or within specific tax jurisdictions.

---

## 2. Key Architecture & Tech Stack

The system is architected as a robust **Single-Page Application (SPA)** that runs entirely client-side, ensuring complete data privacy, high speed, and zero latency.

- **Frontend Core**: Semantic HTML5, Vanilla JavaScript (ES6+ modular structure)
- **Styling**: Vanilla CSS3 with premium design tokens (Glassmorphism, custom dark-mode gradients, fluid typography via standard Google Fonts)
- **Visualization**: Chart.js (via secure CDN) for rendering smooth, animated, interactive financial charts (Income vs Expense trends, Expense categorization, Profit margins)
- **State Management**: Reactive state pattern in JavaScript with automatic persistence to `localStorage`
- **Asset/Data Processing**: High-performance JSON-based import/export engine for record back-ups

---

## 3. Data Dictionary & Schema

The application manages two primary transactional arrays: `income` and `expenses`.

### 3.1 Settings Schema
```typescript
interface SystemSettings {
  currencySymbol: string;     // e.g., "¥", "$", "€", "£"
  currencyCode: string;       // e.g., "JPY", "USD", "EUR", "GBP"
  vatRate: number;            // VAT percentage, e.g., 10 for 10%
  monthlyProfitGoal: number;  // Financial milestone tracking
}
```

### 3.2 Transaction Schema (Income & Expenses)
```typescript
interface Transaction {
  id: string;                 // Unique UUID/timestamp
  type: 'income' | 'expense'; // Transaction category
  date: string;               // ISO date string (YYYY-MM-DD)
  description: string;        // Text summary
  category: string;           // Pre-defined or custom tag
  amountGross: number;        // Total amount including VAT
  amountVat: number;          // Calculated VAT portion: amountGross * (vatRate / (100 + vatRate))
  amountNet: number;          // Net amount: amountGross - amountVat
  isVatApplicable: boolean;   // Whether this item is subject to VAT
}
```

---

## 4. Newly Designed Important Functions

To elevate this from a simple tracker to a high-end **Financial Intelligence System**, we have designed and implemented several core functions:

### 4.1 Real-Time Dynamic Financial Recalculation Engine
*   **Problem**: In typical systems, changing a currency symbol or VAT rate only affects new entries, leaving historical data inconsistent.
*   **Solution**: An observer-pattern-based engine. When the user changes the system currency or VAT rate:
    1.  The system iterates through all existing historical records.
    2.  It dynamically recalculates the VAT portion (`amountVat`) and Net value (`amountNet`) using the new rate for entries marked as VAT-applicable.
    3.  It instantly updates all summaries (Gross, Net, Total Tax, Profit) and redraws charts without requiring page reloads.

### 4.2 Financial Forecasting & Predictive Analysis
*   **Functionality**: Uses linear regression on past historical monthly income and expense data to project financial outcomes for the next 3 months.
*   **Visual Indicator**: Renders a predictive dotted-line projection on the main trend chart, outlining estimated revenue, expenditure, and safety margins.

### 4.3 Interactive Budgeting & Goal Tracker
*   **Functionality**: Allows the company to set monthly gross income and net profit targets.
*   **UI Component**: A gorgeous glassmorphic radial progress tracker and linear bar meters showing real-time progress toward milestones, complete with motivational micro-interactions.

### 4.4 Automated Tax (VAT) Ledger & Compliance Auditor
*   **Functionality**: A dedicated tax center that aggregates VAT collected (from income) vs. VAT paid (from business expenses).
*   **Benefit**: Calculates the net VAT liability/refund due to the tax authority, making corporate filing preparation painless.

### 4.5 Intelligent Financial Advisory System (Rule-Based AI)
*   **Functionality**: Analyzes the company's financial data to provide actionable advice in a beautiful "Insights Console".
*   **Example Insights**:
    *   *“Operating expenses have risen by 15% this month, primarily driven by Software subscriptions. Consider auditing active licenses.”*
    *   *“Your net profit margin is 35%, which exceeds your sector average. It might be time to reinvest surplus in marketing.”*
    *   *“VAT liabilities are accumulating. Ensure at least 1,200,000 ¥ is set aside for the upcoming quarterly filing.”*

---

### 4.6 Profit & Loss (P&L) Statement Generator
*   **Functionality**: Auto-generates a structured monthly / quarterly / yearly P&L statement directly from transactional data.
*   **Sections Computed**:
    *   Gross Revenue (sum of income `amountGross`)
    *   Net Revenue (sum of income `amountNet`, excluding output VAT)
    *   Cost of Goods Sold (expense categories tagged `COGS`)
    *   Gross Profit = Net Revenue − COGS
    *   Operating Expenses (all non-COGS expenses, net of VAT)
    *   Operating Profit (EBIT) = Gross Profit − OpEx
    *   Net Profit Margin (%) = Operating Profit / Net Revenue × 100
*   **Output**: Renders in-app as a polished accounting table, with a one-click export to CSV / printable PDF (¥ formatted with thousand separators, e.g. `¥1,234,567`).

### 4.7 Break-Even Point (BEP) Analyzer
*   **Problem**: SMEs need to know the minimum sales volume that covers all costs.
*   **Functionality**:
    1.  Classifies expenses into **Fixed Costs** (rent, salaries, subscriptions) and **Variable Costs** (materials, commissions) using category tags.
    2.  Computes BEP (in ¥) = Fixed Costs / Contribution Margin Ratio, where Contribution Margin Ratio = (Net Revenue − Variable Costs) / Net Revenue.
    3.  Plots a **break-even chart**: revenue line vs. total cost line, with the intersection point highlighted in neon emerald.
*   **Use**: Helps management decide pricing and minimum monthly sales targets.

### 4.8 Cash Flow Calendar & Liquidity Heatmap
*   **Functionality**: A calendar grid where each day is color-coded by net cash flow (income − expense).
*   **Liquidity Forecast**: Projects running cash balance day-by-day, flagging dates when the balance is predicted to fall below a user-defined **minimum liquidity threshold** (default: `¥500,000`).
*   **Alert**: Triggers an in-app warning banner: *“Projected cash balance falls below threshold on 2026-06-14. Consider deferring expense X or accelerating receivable Y.”*

### 4.9 Multi-Currency FX Conversion Layer
*   **Functionality**: Although JPY (`¥`) is the default base currency, individual transactions may be recorded in USD, EUR, GBP, etc.
*   **Mechanics**:
    *   Each transaction stores an optional `originalCurrency` and `fxRate` (rate to base at transaction date).
    *   All aggregated dashboards always display in the **system base currency** (configurable; default `¥`).
    *   When the user switches base currency (e.g. ¥ → $), the entire dashboard re-converts using the stored FX rates — no data loss.
*   **Benefit**: Useful for SMEs with overseas suppliers or clients.

### 4.10 Category-Level Cost Optimization Recommender
*   **Functionality**: For each expense category, compares the latest month vs. trailing 6-month average and flags categories with the largest cost growth.
*   **Output**: A ranked "Top 5 Cost Pressures" list with suggested savings, e.g. *“Utilities up 28% MoM (avg ¥85,000 → ¥109,000). Estimated annual savings if normalized: ¥288,000.”*

### 4.11 KPI Scorecard Panel
*   **Functionality**: A compact at-a-glance scorecard with the company's seven key financial KPIs:
    1.  Monthly Recurring Revenue (MRR)
    2.  Gross Margin %
    3.  Operating Margin %
    4.  Burn Rate (¥/month)
    5.  Runway (months of cash remaining)
    6.  Average Transaction Value
    7.  Expense-to-Income Ratio
*   **Visual**: Each KPI shows current value, delta vs. previous period (▲ green / ▼ red), and a 12-point sparkline.

### 4.12 VAT Rate History & Retroactive Mode
*   **Problem**: VAT rates can change by law (e.g. Japan 8% → 10%). Retroactively recalculating all historical records using the new rate would be **legally incorrect**.
*   **Solution**: The system stores a **VAT rate timeline** (`{effectiveFrom: "2019-10-01", rate: 10}`, etc.). Each transaction's VAT is calculated using the rate in force on its transaction date. Settings UI lets the user add, edit, and version VAT rates safely.

### 4.13 Audit Log & Data Integrity Hash
*   **Functionality**: Every create / update / delete operation on a transaction is appended to an immutable in-browser audit log with timestamp, action, and a SHA-256 hash chain.
*   **Benefit**: Tamper-evident records — critical for tax audits and internal accountability.

### 4.14 Smart Search & Natural-Language Filter
*   **Functionality**: A command-bar style search that accepts queries like:
    *   `"expenses > 100000 in 2026-Q1 category:software"`
    *   `"income from client:Acme last 3 months"`
*   **Mechanics**: Tokenizes the query, applies range/category/date filters, and instantly updates both the table and the charts.

### 4.15 Snapshot Backup & Restore
*   **Functionality**: One-click export of the entire dataset (transactions, settings, VAT timeline, audit log) into a signed JSON snapshot file. Restore reconstructs the exact application state, enabling safe experimentation and end-of-year archival.

### 4.16 Stock / Inventory Analyzing Module
*   **Purpose**: Track raw materials & finished goods, link them to income/expense transactions, and surface stock health analytics.
*   **Stock Item Schema**:
    ```typescript
    interface StockItem {
      sku: string;                  // Unique stock keeping unit
      name: string;
      category: string;             // e.g. "Raw Material", "Finished Good"
      unit: string;                 // e.g. "pcs", "kg", "box"
      quantityOnHand: number;
      reorderPoint: number;         // Auto-alert when stock falls below
      reorderQuantity: number;      // Suggested PO size
      unitCost: number;             // Weighted average cost (¥, base currency)
      unitPrice: number;            // Selling price (¥)
      lastInboundDate: string;      // ISO date
      lastOutboundDate: string;     // ISO date
      supplierId: string;
    }
    ```
*   **Analytical Functions**:
    1.  **Inventory Valuation** — total value on hand = Σ (`quantityOnHand × unitCost`), shown in ¥ with VAT-exclusive / VAT-inclusive toggle.
    2.  **Stock Turnover Ratio** — COGS / Average Inventory, computed per month and per category. High turnover = healthy sell-through; low turnover = dead stock.
    3.  **Days of Inventory On Hand (DOH)** — `quantityOnHand / averageDailyUsage`, color-coded: green (<30d), amber (30–90d), red (>90d → over-stocked).
    4.  **ABC Classification** — automatic Pareto bucketing: A items (top 70% of value), B (next 20%), C (last 10%). Helps prioritize focus.
    5.  **Slow-Mover / Dead-Stock Detector** — items with zero outbound movement for ≥ 90 days are flagged on a "Dead Stock" card with a write-down suggestion.
    6.  **Reorder Alert Engine** — when `quantityOnHand ≤ reorderPoint`, the system raises a banner and pre-fills a Purchase Order draft (see §4.17).
    7.  **Stock-vs-Sales Correlation Chart** — overlays daily outbound volume on the income trend line so spikes in revenue are tied back to which SKUs sold.
*   **Profit Linkage**: When a sale is recorded, the matched SKU's `unitCost × quantitySold` is automatically booked into COGS, feeding §4.6's Gross Profit calculation — no manual journal entry required.

### 4.17 Purchase Order (PO) Tracking with Status Workflow
*   **Purpose**: Full lifecycle tracking of every supplier purchase, tightly linked to the stock module and the expense ledger.
*   **PO Schema**:
    ```typescript
    interface PurchaseOrder {
      poId: string;                 // e.g. "PO-2026-00042"
      supplierId: string;
      supplierName: string;
      orderDate: string;            // ISO date
      expectedDate: string;         // Promised delivery date
      actualDeliveryDate: string;   // Filled when received
      lineItems: Array<{
        sku: string;
        quantity: number;
        unitCost: number;           // ¥ (in transaction currency)
        vatApplicable: boolean;
      }>;
      currency: string;             // Default "JPY"
      fxRate: number;               // Rate to base currency (¥)
      subtotal: number;             // Σ qty × unitCost
      vatAmount: number;
      totalGross: number;
      status: 'Preparing' | 'Delivery' | 'Not Payment' | 'Done';
      paymentDueDate: string;
      paidDate: string;
      notes: string;
    }
    ```
*   **Status Workflow & Semantics**:
    | Status | Meaning | Triggers | Allowed Next States |
    |---|---|---|---|
    | **Preparing** | PO drafted, supplier confirmation pending, goods not yet shipped. | Created manually or auto-drafted by §4.16 reorder alert. | → Delivery, → Cancelled |
    | **Delivery** | Goods are en route or partially received. | User marks supplier as having shipped; expected date set. | → Not Payment (when fully received) |
    | **Not Payment** | Goods received & added to stock, but invoice **unpaid**. Stock count increases here. | Triggered on "Mark Received". `quantityOnHand` is incremented; expense booked as **Accounts Payable**, not yet cash-out. | → Done |
    | **Done** | Invoice fully paid. Cash outflow recorded in the expense ledger; A/P cleared. | Triggered by "Mark Paid" with payment date + method. | (terminal) |
*   **Visual Workflow**: A horizontal kanban-style strip — *Preparing → Delivery → Not Payment → Done* — with each PO card draggable between columns. Cards show supplier, total ¥, days-in-status, and a status pill (amber for Preparing, blue for Delivery, red for Not Payment, green for Done).
*   **KPIs Surfaced**:
    *   **On-Time Delivery Rate** — % of POs where `actualDeliveryDate ≤ expectedDate`.
    *   **Average Days to Pay** — mean of `paidDate − deliveryDate`; helps optimize working capital.
    *   **Outstanding Payables (¥)** — sum of `totalGross` for all POs in *Not Payment* status; shown on the dashboard's cash-flow card so the user always sees committed-but-unpaid cash obligations.
    *   **Supplier Scorecard** — per-supplier reliability (on-time %, defect rate from received notes, avg lead time).
*   **Overdue Alerts**:
    *   PO stuck in *Delivery* past `expectedDate` → amber warning.
    *   PO stuck in *Not Payment* past `paymentDueDate` → red warning + suggested action *"Pay ¥XXX,XXX to Supplier A by 2026-06-10 to avoid late penalty."*
*   **Tight Integration**:
    *   *Preparing → Delivery*: no financial impact yet.
    *   *Delivery → Not Payment*: stock increases (`StockItem.quantityOnHand`), **Accrued Expense** journal entry created (does not affect cash, but does affect Operating Profit in §4.6).
    *   *Not Payment → Done*: real cash expense booked into the expense ledger with proper VAT split (using the VAT rate from §4.12 effective on `paidDate`).

---

## 5. Website Menu & Navigation Structure

The dashboard is organized as a left-rail navigation menu with the following top-level sections; each opens a dedicated view in the SPA:

| # | Menu Item | Icon | Contents |
|---|-----------|------|----------|
| 1 | **Dashboard** | 🏠 | KPI scorecard (§4.11), income vs expense trend, profit goal radial, quick-add buttons |
| 2 | **Income** | 💰 | Income transactions list, add/edit form, category breakdown chart |
| 3 | **Expenses** | 💳 | Expense transactions list, fixed vs variable tagging, top-5 cost pressures (§4.10) |
| 4 | **Stock / Inventory** | 📦 | Stock list, valuation card, ABC chart, dead-stock alerts, reorder queue (§4.16) |
| 5 | **Purchase Orders** | 🧾 | Kanban board (Preparing / Delivery / Not Payment / Done), supplier scorecard (§4.17) |
| 6 | **Suppliers** | 🏭 | Supplier directory, contact info, payment terms, performance metrics |
| 7 | **Cash Flow** | 📅 | Liquidity heatmap calendar, runway projection (§4.8) |
| 8 | **P&L Statement** | 📊 | Monthly / Quarterly / Yearly P&L, export CSV/PDF (§4.6) |
| 9 | **Break-Even Analysis** | ⚖️ | Fixed/variable split, BEP chart (§4.7) |
| 10 | **VAT / Tax Center** | 🧮 | VAT ledger, rate timeline editor, liability summary (§4.4 + §4.12) |
| 11 | **Forecast & Insights** | 🔮 | 3-month forecasts, AI advisory console (§4.2 + §4.5) |
| 12 | **Reports** | 📁 | Custom report builder, scheduled exports, snapshots (§4.15) |
| 13 | **Audit Log** | 🛡️ | Tamper-evident change history (§4.13) |
| 14 | **Settings** | ⚙️ | Currency (default ¥ JPY), VAT rate (default 10%) + timeline, profit goals, user preferences |

A top bar provides: global search (§4.14), base-currency switcher (¥/$/€/£), current period selector (This Month / Quarter / Year / Custom), and a notification bell for overdue POs and low stock.

---

## 6. Backend: Google Apps Script + Google Sheets as Database

The system is **built on Google Apps Script (GAS)** and uses **Google Sheets as the persistent backend database**, replacing the earlier client-only `localStorage` model. This makes the system multi-user, cloud-synced, and free to host inside any company's Google Workspace.

### 6.1 Why Google Apps Script + Sheets
*   **Zero-cost hosting** — runs entirely on the user's Google account.
*   **Native auth** — Google login handles user identity & permissions.
*   **Familiar database** — accountants can audit raw data directly in Sheets.
*   **Built-in sharing** — share the spreadsheet with the finance team for collaborative access.
*   **Automatic backups** — Google Drive version history acts as a free backup system.

### 6.2 Architecture
```
[ Browser (HTML/CSS/JS frontend) ]
            │
            │  google.script.run.<server function>
            ▼
[ Google Apps Script (Code.gs server-side) ]
            │
            │  SpreadsheetApp.openById(...)
            ▼
[ Google Sheets — multi-tab workbook acting as DB ]
```
*   **Frontend**: served by GAS via `HtmlService.createHtmlOutputFromFile('index')`, deployed as a Web App URL.
*   **Server layer** (`Code.gs`): exposes functions like `getTransactions()`, `addTransaction(obj)`, `updateStock(sku, qty)`, `advancePOStatus(poId, newStatus)`. Called from the browser using `google.script.run.withSuccessHandler(...).serverFn(args)`.
*   **Database**: a single Google Sheets workbook with one tab per entity.

### 6.3 Google Sheets Tab Layout (one tab = one table)
| Tab Name | Purpose | Key Columns |
|---|---|---|
| `Settings` | Single-row config | currencySymbol, currencyCode, vatRate, monthlyProfitGoal, minLiquidity |
| `VATTimeline` | Versioned VAT rates | effectiveFrom, rate, note |
| `Income` | Income transactions | id, date, description, category, amountGross, amountVat, amountNet, isVatApplicable, currency, fxRate |
| `Expenses` | Expense transactions | id, date, description, category, costType (Fixed/Variable/COGS), amountGross, amountVat, amountNet, linkedPOId |
| `Stock` | Inventory items | sku, name, category, unit, quantityOnHand, reorderPoint, reorderQuantity, unitCost, unitPrice, supplierId |
| `StockMoves` | Inbound/outbound history | moveId, sku, date, qtyDelta, reason (PO/Sale/Adjust), refId |
| `PurchaseOrders` | PO headers | poId, supplierId, orderDate, expectedDate, actualDeliveryDate, totalGross, status, paymentDueDate, paidDate |
| `POLineItems` | PO line items | poId, sku, quantity, unitCost, vatApplicable |
| `Suppliers` | Supplier master | supplierId, name, contact, paymentTermsDays |
| `AuditLog` | Tamper-evident change log | timestamp, user, action, entity, entityId, beforeHash, afterHash |

### 6.4 Server API Surface (Apps Script functions)
*   `getBootstrap()` — returns Settings + VATTimeline + recent transactions in one round-trip (minimizes GAS quota usage).
*   `addIncome(tx)` / `addExpense(tx)` — appends a row, recalculates VAT using the rate from `VATTimeline` effective on `tx.date`.
*   `listStock()` / `upsertStockItem(item)` / `adjustStock(sku, delta, reason)`.
*   `createPO(po)` — writes PO header + line items, status = `Preparing`.
*   `advancePOStatus(poId, newStatus)` — enforces the workflow: Preparing → Delivery → Not Payment → Done. On entering *Not Payment*, automatically increments stock and appends an Accrued Expense row. On entering *Done*, appends the paid expense row with correct VAT.
*   `getDashboardSummary(period)` — server-side aggregation for KPIs, P&L, BEP.
*   `exportSnapshot()` / `restoreSnapshot(blob)` — backup/restore from a JSON blob in Drive.

### 6.5 Performance & Quota Considerations
*   GAS executions are rate-limited (~6 min/run, daily quotas). Mitigations:
    *   **Batch reads** with `Range.getValues()` instead of cell-by-cell access.
    *   **CacheService** for `Settings` and `VATTimeline` (5-minute TTL) — these change rarely.
    *   **Client-side caching** of immutable historical months; only the current month is re-fetched on each action.
    *   **PropertiesService** for tiny key/value config (e.g. base currency).
*   For workbooks expected to exceed ~50,000 rows, the design recommends archiving prior-year data into a separate "Archive_YYYY" workbook to keep the live sheet fast.



---

## 7. UI/UX Design Specifications

To ensure the user is "wowed" at first glance, the interface uses a dark-glass dashboard layout:

*   **Color Palette**: Deep Obsidian (`#0F0F1A`), Neon Emerald (`#00F2FE` / `#4FACFE` gradient for Income/Profit), Crimson Sunrise (`#FF416C` / `#FF4B2B` gradient for Expenses), and Soft Amber (`#FCE38A` / `#F38181` for VAT).
*   **Glassmorphism**: Cards styled with `backdrop-filter: blur(20px)` and thin semi-transparent white borders (`rgba(255, 255, 255, 0.08)`).
*   **Interactive State Transitions**: Smooth `transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1)` on hover effects for all controls.
*   **Responsive Grid**: A multi-column dashboard layout that shrinks cleanly down to mobile screen sizes without compromising visual polish.
*   

---

# Adding Funcition

- [ ] Not payment option
  - [ ] Pre-paid 30%/60%
- [ ] Currency Exchanging Tracking 
  - [ ] The Real Profit in YEN/DON
  - [ ] Transaction①: YEN: buy/sell
  - [ ] Transaction②: DON: buy/sell
  - [ ] Transaction③: DON 
  - [ ] Profit in YEN 
