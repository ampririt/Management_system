/**
 * ProfitIntelligence — Google Apps Script Backend
 * --------------------------------------------------------------
 * Frontend served by HtmlService. Persistent data stored in a
 * Google Sheets workbook (one tab per entity).
 *
 * MONEY MODEL (v3):
 *  - Goods have ONE price: unitSell in ¥. VND is always derived = unitSell × qty × fixedRate.
 *    Buying and selling the goods is break-even at the fixed rate by design.
 *  - Create PO  → full ¥ paid immediately from Cash (¥ wallet) or a Credit card.
 *  - PO "Done"  → the Vietnam customer pays in VND → that VND lands in the VND wallet.
 *  - All profit comes from exchanging VND → ¥ at a better-than-fixed rate (Trading & FX).
 *  - Income / Expense are ¥ only. No VAT. No P&L. No inventory/stock.
 */

// =============================================================
// CONFIG & DYNAMIC SYNC
// =============================================================

/** Add a menu to the Google Sheet UI */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 ProfitIntel')
    .addItem('Initialize / Sync Web App', 'setupWorkbook')
    .addToUi();
}

/** Helper to get the Spreadsheet ID dynamically */
function _getSsId() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (id) return id;

  try {
    const activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
    if (activeId) {
      PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", activeId);
      return activeId;
    }
  } catch (e) {}

  throw new Error("Spreadsheet ID not set. Please open your Google Sheet and click 'ProfitIntel' -> 'Initialize' from the top menu.");
}

const TABS = {
  SETTINGS: "Settings",
  INCOME: "Income",
  EXPENSES: "Expenses",
  PURCHASE_ORDERS: "PurchaseOrders",
  PO_LINE_ITEMS: "POLineItems",
  SUPPLIERS: "Suppliers",
  AUDIT_LOG: "AuditLog",
  CONVERSIONS: "Conversions",
  CREDIT_CARDS: "CreditCards",
  CARD_TXNS: "CardTxns",
  RECURRING: "Recurring",
  TRANSACTIONS: "Transactions",
};

const HEADERS = {
  // foreignSymbol/foreignCode = the selling-market currency (default VND). currencySymbol/Code = home/base (default JPY).
  // fixedRateVND = foreign units per 1 base (₫ per ¥). openingCashYen = starting ¥ cash on hand.
  Settings: ["currencySymbol","currencyCode","monthlyIncomeGoal","monthlyProfitGoal","minLiquidity","fixedRateVND","foreignSymbol","foreignCode","openingCashYen"],
  // Personal income/expense — ¥ only.
  Income: ["id","date","description","client","category","amount"],
  Expenses: ["id","date","description","vendor","category","costType","amount","linkedPOId","paymentMethod","cardId"],
  // Buying goods in Japan. totalYen paid up-front from cash/credit; totalVnd is the receivable (totalYen × fixedRate).
  PurchaseOrders: ["poId","supplierId","supplierName","orderDate","expectedDate","actualDeliveryDate","totalYen","totalVnd","status","paymentMethod","cardId","paidDate","notes"],
  // One ¥ price per line: unitSell. VND is derived, never stored as a separate price.
  POLineItems: ["poId","productName","category","quantity","unitSell"],
  Suppliers: ["supplierId","name","contact","paymentTermsDays","rating"],
  AuditLog: ["timestamp","user","action","entity","entityId","beforeHash","afterHash","payload"],
  // VND -> JPY currency conversions (the profit engine). rate = VND per 1 JPY.
  Conversions: ["id","date","vndAmount","rate","jpyReceived","fee","note"],
  // Personal credit cards. Balance is derived from CardTxns + credit POs (opening + charges - payments).
  CreditCards: ["cardId","name","issuer","creditLimit","statementDay","dueDay","openingBalance"],
  // Card charges (deferred, no cash impact) and payments (cash outflow on date).
  CardTxns: ["id","cardId","date","description","category","amount","type"],
  // Recurring auto-charges (subscriptions, utilities). lastPosted = "YYYY-MM" of the last month auto-posted.
  Recurring: ["id","cardId","name","category","amount","dayOfMonth","active","lastPosted"],
  // History of finished sales — feeds the VND wallet. totalVnd = quantity × unitSell × fixedRate.
  Transactions: ["id","date","productName","supplierName","quantity","unitSell","totalVnd","poId"],
};

// =============================================================
// WEB APP ENTRY
// =============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("app")
    .setTitle("ProfitIntelligence — Financial Command Center")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================
// INIT & SYNC — Run from Sheet Menu
// =============================================================
function setupWorkbook() {
  // Resolve the target spreadsheet. From the Sheet menu we have an active spreadsheet;
  // from the deployed web app we don't, so fall back to the stored SPREADSHEET_ID.
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) {
    PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  } else {
    ss = SpreadsheetApp.openById(_getSsId());
  }

  Object.keys(HEADERS).forEach((tabName) => {
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);
    const targetHeaders = HEADERS[tabName];
    if (sh.getLastRow() === 0) {
      sh.appendRow(targetHeaders);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, targetHeaders.length).setFontWeight("bold");
    } else {
      // Migration by NAME: if the live header doesn't match the target order, remap
      // every row into the new column order. Appending blindly (the old approach)
      // corrupted tabs whose columns were inserted in the middle, not at the end.
      const lastCol = Math.max(1, sh.getLastColumn());
      const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      const mismatch = targetHeaders.some((h, i) => existing[i] !== h) || existing.length !== targetHeaders.length;
      if (mismatch) {
        const oldRows = sh.getLastRow() > 1
          ? sh.getRange(2, 1, sh.getLastRow() - 1, existing.length).getValues() : [];
        const remapped = oldRows.map((r) => {
          const o = {}; existing.forEach((h, i) => { o[h] = r[i]; });
          return targetHeaders.map((h) => (o[h] !== undefined ? o[h] : ""));
        });
        sh.clear();
        sh.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]).setFontWeight("bold");
        if (remapped.length) sh.getRange(2, 1, remapped.length, targetHeaders.length).setValues(remapped);
        sh.setFrozenRows(1);
      }
    }
  });

  const settings = ss.getSheetByName(TABS.SETTINGS);
  if (settings.getLastRow() === 1) {
    settings.appendRow(["¥", "JPY", 10000000, 4000000, 500000, 185, "₫", "VND", 0]);
  } else {
    const fillIfBlank = (key, val) => {
      const c = HEADERS.Settings.indexOf(key) + 1;
      if (c <= 0) return;
      const cur = settings.getRange(2, c).getValue();
      if (cur === "" || cur === null || cur === undefined) settings.getRange(2, c).setValue(val);
    };
    fillIfBlank("fixedRateVND", 185);
    fillIfBlank("foreignSymbol", "₫");
    fillIfBlank("foreignCode", "VND");
    fillIfBlank("openingCashYen", 0);
  }

  SpreadsheetApp.flush();
  return "Workbook synced with ID: " + ss.getId();
}

// =============================================================
// LOW-LEVEL HELPERS
// =============================================================
// One Spreadsheet handle per execution (GAS gives each invocation a fresh context),
// so we never pay the openById round-trip more than once per request.
let _SS = null;
function _ss() {
  if (!_SS) _SS = SpreadsheetApp.openById(_getSsId());
  return _SS;
}

function _sheet(name) {
  let sh = _ss().getSheetByName(name);
  if (!sh) {
    setupWorkbook();
    _SS = null;
    sh = _ss().getSheetByName(name);
  }
  return sh;
}

function _rows(sheetName) {
  const sh = _sheet(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = HEADERS[sheetName];
  const values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  return values.map((row) => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) val = val.toISOString();
      else if (val === undefined || val === null) val = "";
      obj[h] = val;
    });
    return obj;
  });
}

// Free-text fields written to a sheet must not begin with a formula trigger, or the
// cell becomes a live formula when the sheet is opened (formula injection).
function _safeText(v) {
  if (v === undefined || v === null) return "";
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function _appendObject(sheetName, obj) {
  const sh = _sheet(sheetName);
  const row = HEADERS[sheetName].map((h) => (obj[h] !== undefined ? obj[h] : ""));
  sh.appendRow(row);
}

// Batched multi-row append — one setValues write instead of N appendRow round-trips.
function _appendObjects(sheetName, objs) {
  if (!objs || !objs.length) return;
  const sh = _sheet(sheetName);
  const headers = HEADERS[sheetName];
  const rows = objs.map((o) => headers.map((h) => (o[h] !== undefined ? o[h] : "")));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function _updateById(sheetName, idKey, idValue, patch) {
  const sh = _sheet(sheetName);
  const headers = HEADERS[sheetName];
  const colIdx = headers.indexOf(idKey);
  const last = sh.getLastRow();
  if (last < 2) return false;
  const data = sh.getRange(2, 1, last - 1, headers.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][colIdx]) === String(idValue)) {
      Object.keys(patch).forEach((k) => {
        const c = headers.indexOf(k);
        if (c >= 0) data[i][c] = patch[k];
      });
      sh.getRange(i + 2, 1, 1, headers.length).setValues([data[i]]);  // single write
      return true;
    }
  }
  return false;
}

function _uuid() { return Utilities.getUuid(); }
function _sha256(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return bytes.map((b) => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

// Serialize every mutation so concurrent writes (incl. postDueRecurring) can't interleave.
function _withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function _audit(action, entity, entityId, before, after) {
  const u = Session.getActiveUser().getEmail() || "unknown";
  _appendObject(TABS.AUDIT_LOG, {
    timestamp: new Date().toISOString(),
    user: u,
    action: action,
    entity: entity,
    entityId: entityId,
    beforeHash: before ? _sha256(JSON.stringify(before)) : "",
    afterHash: after ? _sha256(JSON.stringify(after)) : "",
    payload: JSON.stringify(after || before || {}).slice(0, 500),
  });
}

// Fixed selling rate (foreign units per 1 base, e.g. ₫ per ¥).
function _fixedRate() {
  const s = _rows(TABS.SETTINGS)[0];
  return Number(s && s.fixedRateVND) || 185;
}

// =============================================================
// PUBLIC API
// =============================================================

function getBootstrap() {
  const safe = (label, fn, fallback) => {
    try { return fn(); }
    catch (e) { return fallback; }
  };
  let ssId;
  try {
    ssId = _getSsId();
    const ss = _ss();
    const missing = Object.keys(HEADERS).some((t) => !ss.getSheetByName(t));
    if (missing) { setupWorkbook(); _SS = null; }
  } catch (e) {
    return { ok: false, error: 'Cannot open the database spreadsheet (' + (e.message || e) + '). '
      + 'Open your Google Sheet and run "🚀 ProfitIntel → Initialize / Sync Web App", '
      + 'or set the Script Property SPREADSHEET_ID, then reload.' };
  }
  safe('PostDueRecurring', () => postDueRecurring(), null);  // auto-apply subscriptions/utilities (gated to 1×/day)
  return {
    ok: true,
    spreadsheetId: ssId,
    serverTime: new Date().toISOString(),
    settings:    safe('Settings',    () => _rows(TABS.SETTINGS)[0] || null, null),
    income:      safe('Income',      () => _rows(TABS.INCOME),         []),
    expenses:    safe('Expenses',    () => _rows(TABS.EXPENSES),       []),
    pos:         safe('PurchaseOrders', () => _rows(TABS.PURCHASE_ORDERS), []),
    poLines:     safe('POLineItems', () => _rows(TABS.PO_LINE_ITEMS),  []),
    suppliers:   safe('Suppliers',   () => _rows(TABS.SUPPLIERS),      []),
    conversions: safe('Conversions', () => _rows(TABS.CONVERSIONS),    []),
    creditCards: safe('CreditCards', () => _rows(TABS.CREDIT_CARDS),   []),
    cardTxns:    safe('CardTxns',    () => _rows(TABS.CARD_TXNS),      []),
    recurring:   safe('Recurring',   () => _rows(TABS.RECURRING),      []),
    transactions: safe('Transactions', () => _rows(TABS.TRANSACTIONS),   [])
  };
}

function ping() {
  return { ok: true, spreadsheetId: _getSsId() };
}

/** Check connection and return spreadsheet metadata */
function getSystemStatus() {
  try {
    const id = _getSsId();
    const ss = _ss();
    const tabs = ss.getSheets().map(s => s.getName());
    const missing = Object.keys(HEADERS).filter(h => !tabs.includes(h));
    return {
      ok: true,
      name: ss.getName(),
      id: id,
      url: ss.getUrl(),
      tabs: tabs,
      missingTabs: missing,
      isHealthy: missing.length === 0
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function listAuditLog(limit) {
  const all = _rows(TABS.AUDIT_LOG);
  const sorted = all.slice().sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
  return limit ? sorted.slice(0, Number(limit)) : sorted;
}

function getSettings() { return _rows(TABS.SETTINGS)[0] || null; }
function updateSettings(patch) {
  return _withLock(() => {
    const sh = _sheet(TABS.SETTINGS);
    const headers = HEADERS.Settings;
    if (sh.getLastRow() < 2) sh.appendRow(headers.map((_) => ""));
    Object.keys(patch).forEach((k) => {
      const c = headers.indexOf(k) + 1;
      if (c > 0) sh.getRange(2, c).setValue(patch[k]);
    });
    _audit("UPDATE", "Settings", "singleton", null, patch);
    return getSettings();
  });
}

// =============================================================
// INCOME / EXPENSE — ¥ only
// =============================================================
function addIncome(tx) {
  const id = tx.id || _uuid();
  const row = { id, date: tx.date, description: _safeText(tx.description), client: _safeText(tx.client),
    category: _safeText(tx.category) || "Uncategorized", amount: Number(tx.amount) || 0 };
  _appendObject(TABS.INCOME, row);
  _audit("CREATE", "Income", id, null, row);
  return row;
}

function addExpense(tx) {
  const id = tx.id || _uuid();
  const row = { id, date: tx.date, description: _safeText(tx.description), vendor: _safeText(tx.vendor),
    category: _safeText(tx.category) || "Uncategorized", costType: tx.costType || "Variable",
    amount: Number(tx.amount) || 0, linkedPOId: tx.linkedPOId || "",
    paymentMethod: tx.paymentMethod === "credit" ? "credit" : "cash", cardId: tx.cardId || "" };
  _appendObject(TABS.EXPENSES, row);
  _audit("CREATE", "Expense", id, null, row);
  return row;
}

function deleteTransaction(kind, id) {
  const tabName = kind === "income" ? TABS.INCOME : TABS.EXPENSES;
  return _withLock(() => {
    const sh = _sheet(tabName);
    const last = sh.getLastRow();
    if (last < 2) return false;
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.deleteRow(i + 2);
        _audit("DELETE", kind, id, { id }, null);
        return true;
      }
    }
    return false;
  });
}

function listIncome() { return _rows(TABS.INCOME); }
function listExpenses() { return _rows(TABS.EXPENSES); }
function listTransactions() { return _rows(TABS.TRANSACTIONS); }

// =============================================================
// PURCHASE ORDERS — buy ¥ at creation, receive VND at "Done"
// =============================================================
const PO_FLOW = { 'Preparing':'Delivery', 'Delivery':'Not Payment', 'Not Payment':'Done', 'Done':null };

function listPurchaseOrders() {
  const headers = _rows(TABS.PURCHASE_ORDERS);
  const lines = _rows(TABS.PO_LINE_ITEMS);
  return headers.map((po) => Object.assign({}, po, { lineItems: lines.filter((l) => l.poId === po.poId) }));
}

function createPO(po) {
  return _withLock(() => {
    const fx = _fixedRate();
    const poId = po.poId || "PO-" + new Date().getFullYear() + "-" + String(Date.now()).slice(-5);
    const orderDate = po.orderDate || new Date().toISOString().slice(0, 10);
    let totalYen = 0;
    (po.lineItems || []).forEach((li) => { totalYen += Number(li.quantity) * Number(li.unitSell); });
    totalYen = Math.round(totalYen);
    const header = {
      poId, supplierId: po.supplierId || "", supplierName: _safeText(po.supplierName),
      orderDate, expectedDate: po.expectedDate || "", actualDeliveryDate: "",
      totalYen, totalVnd: Math.round(totalYen * fx), status: "Preparing",
      paymentMethod: po.paymentMethod === "credit" ? "credit" : "cash", cardId: po.cardId || "",
      paidDate: "", notes: _safeText(po.notes)
    };
    _appendObject(TABS.PURCHASE_ORDERS, header);
    _appendObjects(TABS.PO_LINE_ITEMS, (po.lineItems || []).map((li) => ({
      poId, productName: _safeText(li.productName), category: _safeText(li.category),
      quantity: Number(li.quantity) || 0, unitSell: Number(li.unitSell) || 0
    })));
    _audit("CREATE", "PurchaseOrder", poId, null, header);
    return header;
  });
}

function updatePO(po) {
  return _withLock(() => {
    if (po.poId) {
      _deleteRow(TABS.PURCHASE_ORDERS, "poId", po.poId);
      const sh = _sheet(TABS.PO_LINE_ITEMS); const last = sh.getLastRow();
      if (last >= 2) {
        const ids = sh.getRange(2, 1, last - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(po.poId)) sh.deleteRow(i + 2); }
      }
    }
    return createPO(po);
  });
}

function advancePOStatus(poId, newStatus, opts) {
  opts = opts || {};
  return _withLock(() => {
    const pos = _rows(TABS.PURCHASE_ORDERS);
    const po = pos.find((p) => p.poId === poId);
    if (!po) throw new Error("PO not found: " + poId);
    const fx = _fixedRate();
    const patch = { status: newStatus };
    if (newStatus === "Not Payment") {
      // Goods received in Vietnam, awaiting the customer's VND payment.
      patch.actualDeliveryDate = opts.actualDeliveryDate || new Date().toISOString().slice(0, 10);
    }
    if (newStatus === "Done") {
      // Customer paid → VND received into the VND wallet (recorded as Transactions rows).
      patch.paidDate = opts.paidDate || new Date().toISOString().slice(0, 10);
      const lines = _rows(TABS.PO_LINE_ITEMS).filter((l) => l.poId === poId);
      _appendObjects(TABS.TRANSACTIONS, lines.map((l) => {
        const qty = Number(l.quantity) || 0;
        const uSell = Number(l.unitSell) || 0;
        return {
          id: _uuid(), date: patch.paidDate, productName: l.productName, supplierName: po.supplierName,
          quantity: qty, unitSell: uSell, totalVnd: Math.round(qty * uSell * fx), poId: poId
        };
      }));
    }
    _updateById(TABS.PURCHASE_ORDERS, "poId", poId, patch);
    _audit("UPDATE", "PurchaseOrder", poId, { status: po.status }, patch);
    return Object.assign({}, po, patch);
  });
}

function listSuppliers() { return _rows(TABS.SUPPLIERS); }
function upsertSupplier(s) {
  return _withLock(() => {
    const existing = _rows(TABS.SUPPLIERS).find((x) => x.supplierId === s.supplierId);
    if (existing) { _updateById(TABS.SUPPLIERS, "supplierId", s.supplierId, s); _audit("UPDATE", "Supplier", s.supplierId, existing, s); }
    else { _appendObject(TABS.SUPPLIERS, s); _audit("CREATE", "Supplier", s.supplierId, null, s); }
    return s;
  });
}

// =============================================================
// CURRENCY CONVERSIONS (VND -> JPY) — the profit engine
// =============================================================
function listConversions() { return _rows(TABS.CONVERSIONS); }
function addConversion(tx) {
  const id = tx.id || _uuid();
  const vndAmount = Number(tx.vndAmount) || 0;
  const rate = Number(tx.rate) || 0;            // VND per 1 JPY
  const fee = Number(tx.fee) || 0;              // JPY fee
  const jpyReceived = rate > 0 ? Math.round(vndAmount / rate) : 0;
  const row = { id, date: tx.date, vndAmount, rate, jpyReceived, fee, note: _safeText(tx.note) };
  _appendObject(TABS.CONVERSIONS, row);
  _audit("CREATE", "Conversion", id, null, row);
  return row;
}
function saveConversion(tx) { return _withLock(() => { if (tx.id) _deleteRow(TABS.CONVERSIONS, "id", tx.id); return addConversion(tx); }); }
function deleteConversion(id) { return _withLock(() => { _audit("DELETE", "Conversion", id, { id }, null); return _deleteRow(TABS.CONVERSIONS, "id", id); }); }

// =============================================================
// CREDIT CARDS (personal cash-flow)
// =============================================================
function listCreditCards() { return _rows(TABS.CREDIT_CARDS); }
function upsertCreditCard(c) {
  return _withLock(() => {
    const cardId = c.cardId || _uuid();
    const row = { cardId, name: _safeText(c.name) || "Card", issuer: _safeText(c.issuer), creditLimit: Number(c.creditLimit) || 0, statementDay: Number(c.statementDay) || 1, dueDay: Number(c.dueDay) || 10, openingBalance: Number(c.openingBalance) || 0 };
    const existing = _rows(TABS.CREDIT_CARDS).find((x) => x.cardId === cardId);
    if (existing) { _updateById(TABS.CREDIT_CARDS, "cardId", cardId, row); _audit("UPDATE", "CreditCard", cardId, existing, row); }
    else { _appendObject(TABS.CREDIT_CARDS, row); _audit("CREATE", "CreditCard", cardId, null, row); }
    return row;
  });
}
function deleteCreditCard(cardId) {
  return _withLock(() => {
    _deleteRow(TABS.CREDIT_CARDS, "cardId", cardId);
    const sh = _sheet(TABS.CARD_TXNS); const last = sh.getLastRow();
    if (last >= 2) { const col = HEADERS.CardTxns.indexOf("cardId") + 1; const ids = sh.getRange(2, col, last - 1, 1).getValues(); for (let i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(cardId)) sh.deleteRow(i + 2); } }
    _audit("DELETE", "CreditCard", cardId, { cardId }, null);
    return true;
  });
}
function listCardTxns() { return _rows(TABS.CARD_TXNS); }
function addCardTxn(tx) {
  const id = tx.id || _uuid();
  const row = { id, cardId: tx.cardId, date: tx.date, description: _safeText(tx.description), category: _safeText(tx.category) || "Other", amount: Number(tx.amount) || 0, type: tx.type === "payment" ? "payment" : "charge" };
  _appendObject(TABS.CARD_TXNS, row);
  _audit("CREATE", "CardTxn", id, null, row);
  return row;
}
function saveCardTxn(tx) { return _withLock(() => { if (tx.id) _deleteRow(TABS.CARD_TXNS, "id", tx.id); return addCardTxn(tx); }); }
function deleteCardTxn(id) { return _withLock(() => { _audit("DELETE", "CardTxn", id, { id }, null); return _deleteRow(TABS.CARD_TXNS, "id", id); }); }

// =============================================================
// RECURRING AUTO-CHARGES (subscriptions, utilities)
// =============================================================
function listRecurring() { return _rows(TABS.RECURRING); }
function upsertRecurring(r) {
  return _withLock(() => {
    const id = r.id || _uuid();
    const row = { id, cardId: r.cardId || "", name: _safeText(r.name) || "Subscription", category: _safeText(r.category) || "Subscription",
      amount: Number(r.amount) || 0, dayOfMonth: Math.min(28, Math.max(1, Number(r.dayOfMonth) || 1)),
      active: r.active === false ? false : true, lastPosted: r.lastPosted || "" };
    const existing = _rows(TABS.RECURRING).find((x) => x.id === id);
    if (existing) { _updateById(TABS.RECURRING, "id", id, row); _audit("UPDATE", "Recurring", id, existing, row); }
    else { _appendObject(TABS.RECURRING, row); _audit("CREATE", "Recurring", id, null, row); }
    return row;
  });
}
function deleteRecurring(id) { return _withLock(() => { _audit("DELETE", "Recurring", id, { id }, null); return _deleteRow(TABS.RECURRING, "id", id); }); }

/**
 * Post any due recurring charges as card transactions. Idempotent: each rule
 * tracks lastPosted ("YYYY-MM"). Gated to run at most once per day, and serialized
 * under the script lock so a refresh racing a user write can't double-post.
 */
function postDueRecurring() {
  return _withLock(() => {
    const props = PropertiesService.getScriptProperties();
    const today = new Date().toISOString().slice(0, 10);
    if (props.getProperty("LAST_RECURRING_RUN") === today) return { posted: 0, skipped: true };

    const rules = _rows(TABS.RECURRING);
    const now = new Date();
    const curYM = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2);  // local "YYYY-MM"
    const curDay = now.getDate();
    let posted = 0;
    rules.forEach((r) => {
      if (r.active === false || String(r.active).toLowerCase() === "false") return;
      if (!r.cardId || !(Number(r.amount) > 0)) return;
      const day = Math.min(28, Math.max(1, Number(r.dayOfMonth) || 1));
      const start = r.lastPosted ? _nextYM(String(r.lastPosted).slice(0, 7)) : curYM;
      let ym = start;
      let guard = 0;
      while (ym <= curYM && guard++ < 120) {
        const isCurrent = (ym === curYM);
        if (!isCurrent || curDay >= day) {
          const dateStr = ym + "-" + String(day).padStart(2, "0");
          // Mark posted BEFORE charging so a mid-loop failure can't double-charge.
          _updateById(TABS.RECURRING, "id", r.id, { lastPosted: ym });
          addCardTxn({ cardId: r.cardId, date: dateStr, type: "charge", amount: Number(r.amount),
            category: r.category || "Subscription", description: "[Auto] " + (r.name || "Recurring") });
          posted++;
        }
        if (isCurrent) break;
        ym = _nextYM(ym);
      }
    });
    props.setProperty("LAST_RECURRING_RUN", today);
    return { posted };
  });
}
// Advance a "YYYY-MM" string by one month using integer math only (timezone-safe).
function _nextYM(ym) {
  let y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7)); // m is 1-based
  m += 1; if (m > 12) { m = 1; y += 1; }
  return y + "-" + ("0" + m).slice(-2);
}

function exportSnapshot() {
  const data = { exportedAt: new Date().toISOString(), settings: _rows(TABS.SETTINGS), income: _rows(TABS.INCOME), expenses: _rows(TABS.EXPENSES), pos: _rows(TABS.PURCHASE_ORDERS), poLines: _rows(TABS.PO_LINE_ITEMS), suppliers: _rows(TABS.SUPPLIERS), conversions: _rows(TABS.CONVERSIONS), creditCards: _rows(TABS.CREDIT_CARDS), cardTxns: _rows(TABS.CARD_TXNS), recurring: _rows(TABS.RECURRING), auditLog: _rows(TABS.AUDIT_LOG), transactions: _rows(TABS.TRANSACTIONS) };
  data.signature = _sha256(JSON.stringify(data));
  return data;
}

function seedDemoData() {
  try { resetAllData(); } catch (e) {}
  const dAgo = (n) => { const d = new Date(); d.setDate(d.getDate() - n); return d.toISOString().slice(0, 10); };
  const mAgo = (m, day) => { const d = new Date(); d.setMonth(d.getMonth() - m); d.setDate(day); return d.toISOString().slice(0, 10); };
  const today = dAgo(0);

  updateSettings({ openingCashYen: 800000, fixedRateVND: 185 });

  // Japanese suppliers
  [{supplierId:"SUP-001",name:"Yodobashi Camera",contact:"wholesale@yodobashi.jp",paymentTermsDays:0,rating:"A+"},
   {supplierId:"SUP-002",name:"Mercari Japan",contact:"orders@mercari.jp",paymentTermsDays:0,rating:"A"},
   {supplierId:"SUP-003",name:"Don Quijote",contact:"b2b@donki.jp",paymentTermsDays:15,rating:"B"}].forEach(upsertSupplier);

  // Credit card
  const card = upsertCreditCard({name:"Rakuten Card", issuer:"Rakuten", creditLimit:500000, statementDay:27, dueDay:10, openingBalance:0});

  // Personal income (¥) across 3 months
  addIncome({date:mAgo(2,25),description:"Salary",client:"Day job",category:"Services",amount:280000});
  addIncome({date:mAgo(1,25),description:"Salary",client:"Day job",category:"Services",amount:280000});
  addIncome({date:mAgo(0,25),description:"Salary",client:"Day job",category:"Services",amount:280000});
  addIncome({date:dAgo(12),description:"Freelance design",client:"Acme",category:"Services",amount:95000});

  // Personal expenses (¥) — some cash, one on credit
  addExpense({date:mAgo(2,1),description:"Rent",vendor:"Landlord",category:"Rent",costType:"Fixed",amount:85000,paymentMethod:"cash"});
  addExpense({date:mAgo(1,1),description:"Rent",vendor:"Landlord",category:"Rent",costType:"Fixed",amount:85000,paymentMethod:"cash"});
  addExpense({date:mAgo(0,1),description:"Rent",vendor:"Landlord",category:"Rent",costType:"Fixed",amount:85000,paymentMethod:"cash"});
  addExpense({date:dAgo(9),description:"New laptop (on card)",vendor:"BIC Camera",category:"Software",costType:"Fixed",amount:140000,paymentMethod:"credit",cardId:card.cardId});
  addExpense({date:dAgo(5),description:"Groceries",vendor:"Supermarket",category:"Other",costType:"Variable",amount:18500,paymentMethod:"cash"});

  // VND -> JPY conversions (the profit). Fixed baseline is 185 ₫/¥; lower real rate = profit.
  addConversion({date:mAgo(1,15), vndAmount:12000000, rate:179, fee:3000, note:"VND strong — good rate"});
  addConversion({date:dAgo(8), vndAmount:7000000, rate:176, fee:2000, note:"Best rate so far"});

  // Credit card extras + recurring bills
  addCardTxn({cardId:card.cardId, date:dAgo(40), type:"charge", amount:23000, category:"Dining", description:"Restaurant"});
  addCardTxn({cardId:card.cardId, date:dAgo(15), type:"payment", amount:50000, category:"Other", description:"Card payment"});
  upsertRecurring({cardId:card.cardId, name:"Netflix", category:"Subscription", amount:1980, dayOfMonth:5, active:true});
  upsertRecurring({cardId:card.cardId, name:"Electricity Bill", category:"Utilities", amount:8500, dayOfMonth:25, active:true});
  PropertiesService.getScriptProperties().deleteProperty("LAST_RECURRING_RUN");
  postDueRecurring();

  // A finished PO (¥ paid on cash, customer's VND received) → builds the VND wallet
  var donePO = createPO({supplierId:"SUP-002", supplierName:"Mercari Japan", orderDate:dAgo(40), expectedDate:dAgo(30),
    paymentMethod:"cash",
    lineItems:[{productName:"Vintage Seiko Watch", category:"Finished Good", quantity:2, unitSell:27000}]});
  advancePOStatus(donePO.poId, "Delivery", {});
  advancePOStatus(donePO.poId, "Not Payment", {});
  advancePOStatus(donePO.poId, "Done", {paidDate: dAgo(10)});

  // An in-flight PO (¥ paid on credit, goods received, awaiting the customer's VND)
  var openPO = createPO({supplierId:"SUP-001", supplierName:"Yodobashi Camera", orderDate:dAgo(18), expectedDate:dAgo(8),
    paymentMethod:"credit", cardId:card.cardId,
    lineItems:[{productName:"Used Mirrorless Camera", category:"Finished Good", quantity:5, unitSell:60000}]});
  advancePOStatus(openPO.poId, "Delivery", {});
  advancePOStatus(openPO.poId, "Not Payment", {});
  return "Demo data seeded.";
}

function saveIncome(tx) { return _withLock(() => { if (tx.id) _deleteRow(TABS.INCOME, "id", tx.id); return addIncome(tx); }); }
function saveExpense(tx) { return _withLock(() => { if (tx.id) _deleteRow(TABS.EXPENSES, "id", tx.id); return addExpense(tx); }); }
function deleteIncome(id) { return _withLock(() => { _audit("DELETE", "Income", id, { id }, null); return _deleteRow(TABS.INCOME, "id", id); }); }
function deleteExpense(id) { return _withLock(() => { _audit("DELETE", "Expense", id, { id }, null); return _deleteRow(TABS.EXPENSES, "id", id); }); }
function deleteSupplier(supplierId) { return _withLock(() => { _audit("DELETE", "Supplier", supplierId, { supplierId }, null); return _deleteRow(TABS.SUPPLIERS, "supplierId", supplierId); }); }
function deletePO(poId) {
  return _withLock(() => {
    _deleteRow(TABS.PURCHASE_ORDERS, "poId", poId);
    const sh = _sheet(TABS.PO_LINE_ITEMS); const last = sh.getLastRow();
    if (last >= 2) { const ids = sh.getRange(2, 1, last - 1, 1).getValues(); for (let i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(poId)) sh.deleteRow(i + 2); } }
    _audit("DELETE", "PurchaseOrder", poId, { poId }, null); return true;
  });
}

function resetAllData() {
  return _withLock(() => {
    ["Income","Expenses","PurchaseOrders","POLineItems","Suppliers","Conversions","CreditCards","CardTxns","Recurring","Transactions"].forEach((tabName) => {
      try {
        const sh = _sheet(tabName);
        const last = sh.getLastRow();
        const cols = Math.max(1, sh.getLastColumn());
        if (last > 1) sh.getRange(2, 1, last - 1, cols).clearContent();
      } catch (e) { /* skip a tab that can't be cleared */ }
    });
    _audit("DELETE", "ALL", "*", null, null);
    return "All transactional data cleared.";
  });
}

function _deleteRow(sheetName, idKey, idValue) { const sh = _sheet(sheetName); const headers = HEADERS[sheetName]; const colIdx = headers.indexOf(idKey) + 1; const last = sh.getLastRow(); if (last < 2) return false; const ids = sh.getRange(2, colIdx, last - 1, 1).getValues(); for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(idValue)) { sh.deleteRow(i + 2); return true; } } return false; }
