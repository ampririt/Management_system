# ProfitIntelligence — Google Apps Script Deployment

Standalone GAS web app backed by a Google Sheets database.

## Files in this folder

| File | Type in GAS editor |
|---|---|
| `appsscript.json` | Manifest (enable **Project Settings → Show "appsscript.json"**) |
| `Code.gs` | Apps Script · server-side |
| `Index.html` | HTML · main UI (template) |
| `Style.html` | HTML · CSS (included by Index) |
| `JavaScript.html` | HTML · client JS (included by Index) |

## One-time setup

1. **Create a Google Sheet** — open <https://sheets.new>, name it e.g. `ProfitIntel DB`. Copy its **Spreadsheet ID** from the URL (the long token between `/d/` and `/edit`).

2. **Create an Apps Script project** — open <https://script.google.com> → **New project**.

3. **Add the files** — in the GAS editor:
   - Rename the default `Code.gs` and paste in the contents of `Code.gs`.
   - Click **+ → HTML** three times and create `Index`, `Style`, `JavaScript`. Paste each file's contents (without the filename extension in the editor — GAS adds `.html` automatically).
   - In **Project Settings**, tick **Show "appsscript.json" manifest file in editor**, then open `appsscript.json` in the editor and replace its contents with the version in this folder.

4. **Set the spreadsheet ID** — in the GAS editor:
   - **Project Settings → Script Properties → Add script property**
   - Key: `SPREADSHEET_ID`  Value: *your spreadsheet ID*
   - Save.

5. **Initialize the workbook** — in the editor, select the function `setupWorkbook` from the dropdown next to the Run button, then click **Run**. You'll be asked to authorize the requested scopes (Sheets + UI + external requests). Approve.

6. **(Optional) Seed demo data** — select `seedDemoData` and click **Run**. This populates suppliers, stock, income, expenses, and POs that walk through every status (Preparing → Delivery → Not Payment → Done).

7. **Deploy as Web App**:
   - **Deploy → New deployment → Web app**
   - Description: `ProfitIntel v1`
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone within your domain** (or as needed)
   - Click **Deploy**, copy the web app URL, open it.

## Architecture summary

```
[ Browser ] ──google.script.run──> [ Code.gs ] ──SpreadsheetApp──> [ Sheets DB ]
```

| Sheet tab | Stores |
|---|---|
| `Settings` | currency, VAT rate, monthly goals, min liquidity |
| `VATTimeline` | versioned VAT rates (retroactive-safe) |
| `Income` | income transactions |
| `Expenses` | expense transactions |
| `Stock` | inventory items |
| `StockMoves` | inbound/outbound stock history |
| `PurchaseOrders` | PO headers + status |
| `POLineItems` | PO line items |
| `Suppliers` | supplier master |
| `AuditLog` | tamper-evident SHA-256 hash chain |

## PO status workflow

`Preparing → Delivery → Not Payment → Done` is enforced by `advancePOStatus(poId, newStatus)`:

- **Delivery**: no side effect (record only)
- **Not Payment**: stock incremented for every line item; an accrued expense row is created (affects EBIT, not cash)
- **Done**: a paid expense row is created with correct VAT split using the rate from `VATTimeline` effective on `paidDate`

## Performance notes

- `getBootstrap()` returns Settings + VATTimeline + ledger + stock + POs in **one** round-trip.
- `_vatRateForDate()` caches the timeline in `CacheService` for 5 minutes.
- All Sheets reads use `Range.getValues()` (batched), never cell-by-cell.

## Security

- Web app deployed as **"User accessing the web app"** — every `AuditLog` row records the acting user's email.
- Data never leaves your Google Workspace tenant.
- Snapshots include a SHA-256 signature for tamper detection.

## Updating

Re-deploy via **Deploy → Manage deployments → ✏️ Edit → New version**. The web app URL stays the same.
