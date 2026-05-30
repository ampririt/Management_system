# ProfitIntelligence — Project Instructions

This project is a financial management system designed for tracking personal finance and cross-border trading (specifically JPY/VND). It utilizes Google Apps Script (GAS) as the backend and Google Sheets as the persistent database.

## 🛠 Tech Stack
- **Backend:** Google Apps Script (`Code.gs`)
- **Database:** Google Sheets (One tab per entity)
- **Frontend:** Single Page Application (SPA) in `app.html`
- **UI Libraries:** Vanilla CSS, Material Icons, Chart.js

## 🏗 Architecture & Design Patterns

### Backend (GAS)
- **Schema Definition:** All sheet names and column headers are centrally managed in `TABS` and `HEADERS` constants in `Code.gs`.
- **Helper Functions:** Low-level database operations are prefixed with an underscore (e.g., `_sheet`, `_rows`, `_appendObject`) to distinguish them from public API functions.
- **Initialization:** The `setupWorkbook` and `getBootstrap` functions handle schema initialization and data hydration for the frontend.
- **Audit Logging:** Every mutation (Create/Update/Delete) is logged in the `AuditLog` sheet with a SHA-256 hash of the payload for integrity.

### Frontend (SPA)
- **State Management:** A global `STATE` object holds the current data. It is hydrated from the server via `getBootstrap`.
- **Server Communication:** All calls to the GAS backend must use the `srv(functionName, args, onDone)` wrapper. This wrapper handles:
    - `google.script.run` execution.
    - Loading/Sync status indicators.
    - **Offline Demo Mode:** If the app is opened outside the GAS environment, it automatically switches to a mock server logic using `buildMock()` and `demoSrv()`.
- **Styling:** Follow the CSS variable system defined in `:root` for consistency (colors, shadows, radius).

## 🔄 Core Workflows

### Purchase Orders (PO)
Used for Japan-side procurement.
`Preparing` ➔ `Delivery` ➔ `Not Payment` (Goods Received) ➔ `Done` (Paid)

### Sales (SO)
Used for Vietnam-side customer orders.
`Preparing` ➔ `Delivery` (Stock Out) ➔ `Awaiting Payment` ➔ `Done`

### Treasury & FX
- **Base Currency:** JPY (¥)
- **Selling Currency:** VND (₫)
- **Fixed Rate:** A locked rate (default 185) used for baseline profit calculations. Realized FX gains/losses are calculated whenever `Conversions` are recorded.

## 📏 Coding Conventions
- **Naming:** Use camelCase for functions and variables.
- **Date Format:** Always use ISO `YYYY-MM-DD` strings for dates to ensure cross-timezone consistency.
- **UI:** Maintain the "Glassmorphism-lite" aesthetic (soft shadows, high contrast text on light surfaces).
- **Security:** Never hardcode Spreadsheet IDs. Use `PropertiesService` or `SpreadsheetApp.getActiveSpreadsheet().getId()`.

## 🚨 Strict Preservation Rule
- **Surgical Edits Only:** When modifying files, NEVER delete or alter parts of the code that are unrelated to the current task or instruction.
- **Code Integrity:** Ensure that existing logic, helpers, and UI elements are preserved exactly as they are unless explicitly ordered to change or remove them.
