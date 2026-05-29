# Finance & Trading Dashboard

A personal-finance and **international-trading** dashboard. It's built for a simple
arbitrage business — **buy products in Japan (pay in ¥), resell in Vietnam (receive ₫),
then convert ₫ back to ¥** — and reports everything as **profit in your home currency (YEN)**.
It also tracks normal life: income, expenses, credit cards, and cash flow.

Both the home and foreign currencies are configurable in **Settings**, so it works for any
"buy in currency A, sell in currency B at a fixed rate" setup — not just ¥/₫.

---

## Two ways to run it

### A. Quick preview (no setup) — *offline demo mode*
Just open **`app.html`** in any browser (double-click it, or drag it into Chrome).
When it can't find the Google backend it automatically loads a **rich sample dataset**
(several months of sales, conversions, cards, subscriptions and orders) and runs fully
in memory. You can click around, add/edit/delete records, change currencies, etc.
Changes last until you reload the page. This is the fastest way to *see how it works*.

> A small "Demo mode" toast appears on load to confirm you're on sample data.

### B. Full deployment — Google Apps Script + Google Sheets (data is saved)
1. **Create a Google Sheet** (<https://sheets.new>) — this becomes the database.
2. **Create an Apps Script project** (<https://script.google.com> → New project).
3. **Add the code:**
   - Paste `Code.gs` into the script editor's `Code.gs`.
   - Click **＋ → HTML**, name it **`app`** (exactly), and paste in `app.html`.
4. **Link the sheet:** in the Sheet, run the menu **🚀 ProfitIntel → Initialize / Sync Web App**
   (or run `setupWorkbook` once from the editor). This creates all tabs and stores the
   spreadsheet ID. Approve the authorization prompt.
5. **(Optional) demo data:** run `seedDemoData`, or use **Settings → Reload Demo Data** in the app.
6. **Deploy → New deployment → Web app**, *Execute as: Me*, then open the web-app URL.

To update later: **Deploy → Manage deployments → Edit → New version** (URL stays the same).

---

## How profit is calculated (the important part)

There are **two exchange rates** and keeping them apart is the whole idea:

1. **Fixed selling rate** (e.g. `1 ¥ = 185 ₫`, set in Settings) — a *pricing* decision you lock in.
2. **Actual conversion rate** — the real market rate the day you swap ₫ → ¥. It moves, and
   you can hold ₫ and wait for a good rate.

Profit (always shown in your **home/base** currency) breaks into three layers:

| Layer | Meaning |
|---|---|
| **Trading margin @fixed** = `sales÷185 − COGS` | profit from your markup, valued at the fixed rate |
| **Realized FX** = `Σ(¥ received − ₫ converted ÷ 185)` | gain/loss from *timing* your conversions (rate **below** 185 = gain) |
| **Unrealized FX** = `held ₫ × (1/rate_now − 1/185)` | paper gain on ₫ still in your wallet |

- **Realized net profit** = Trading margin + Realized FX
- **Economic profit** = Realized net profit + Unrealized FX (paper gain on held ₫)

The **₫ wallet** is a running balance: every sale adds ₫, every conversion removes ₫.
On the Trading page, set the **current market rate** to instantly see *"if I convert the whole
wallet now, I'd receive X, total profit Y, margin Z%"*.

---

## Pages

| Page | What it does |
|---|---|
| **Home** | Cash on hand, credit available/used, trading net profit, ₫ wallet, income/expenses, best sellers, trend & expense charts |
| **Income / Expenses** | Everyday personal ledgers |
| **Credit Cards** | Cards with limits & balances; **charges are deferred** (no cash impact), **payments are cash outflows**. Includes **Subscriptions & Auto-Charges** (electricity, Netflix…) that auto-post each month |
| **Trading & FX** | The profit model above: revenue/COGS/margin/FX, ₫ wallet, conversion log, rate-vs-fixed chart, profit breakdown, and the *convert-now* projection |
| **Stock** | Products. Add a product by entering **cost (¥)** and **sell price (¥)** — the **₫ price is auto-calculated** = ¥ × fixed rate |
| **Purchase Orders** | Workflow board: *Preparing → Delivery → Awaiting Payment → Done*. Supports **deposit + balance** (e.g. 30% prepaid, 70% on receipt) with a paid/owed progress bar. Columns with > 8 orders collapse with a **Show all** toggle |
| **Suppliers** | Vendor directory |
| **Cash Flow** | Inflow/outflow/net KPIs, running-balance line, 6-month bars, and a daily heatmap |
| **P&L** | Auto-generated profit & loss from the ledger |
| **Settings** | Base & selling **currencies**, the **fixed rate**, VAT, export, demo data, reset |

---

## Common tasks

- **Add a product:** Stock → Add Product → enter ¥ cost and ¥ sell price → the ₫ price fills in automatically.
- **Record a sale:** on Stock, lower a product's quantity (the `−` button). Outbound moves feed ₫ revenue, best sellers, and trading metrics.
- **Log a currency exchange:** Trading & FX → *Convert ₫ → ¥* → enter the ₫ amount, the real rate, and any fee. The log shows gain/loss vs the fixed rate.
- **Buy something on credit:** Credit Cards → Add Charge (this raises the card balance but doesn't touch cash). Pay it later with a *Payment* (cash outflow → shows on Cash Flow).
- **Set up a bill/subscription:** Credit Cards → Add Auto-Charge (amount + day of month). It posts itself each month; use **Run now** to catch up.
- **Track a customer deposit:** Purchase Orders → set a Deposit on creation, then **Pay** to record the balance; the card shows paid vs owed and flags **PAID** when settled.
- **Change currency:** Settings → set base/selling symbols & codes and the fixed rate → Save.

---

## Data model (Google Sheets tabs)

`Settings`, `Income`, `Expenses`, `Stock`, `StockMoves`, `PurchaseOrders`, `POLineItems`,
`Suppliers`, `Conversions` (₫→¥ exchanges), `CreditCards`, `CardTxns`, `Recurring`
(auto-charges), `VATTimeline`, `AuditLog`.

```
[ Browser: app.html ] ──google.script.run──> [ Code.gs ] ──SpreadsheetApp──> [ Google Sheets ]
                       (offline demo mode if no backend)
```

`getBootstrap()` returns the whole dataset in one round-trip and auto-posts any due
recurring charges. Reads are batched via `Range.getValues()`.
