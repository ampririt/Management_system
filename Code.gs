/**
 * ProfitIntelligence — Google Apps Script Backend
 * --------------------------------------------------------------
 * Frontend served by HtmlService. Persistent data stored in a
 * Google Sheets workbook (one tab per entity).
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
  VAT_TIMELINE: "VATTimeline",
  INCOME: "Income",
  EXPENSES: "Expenses",
  STOCK: "Stock",
  STOCK_MOVES: "StockMoves",
  PURCHASE_ORDERS: "PurchaseOrders",
  PO_LINE_ITEMS: "POLineItems",
  SUPPLIERS: "Suppliers",
  AUDIT_LOG: "AuditLog",
  CONVERSIONS: "Conversions",
  CREDIT_CARDS: "CreditCards",
  CARD_TXNS: "CardTxns",
  RECURRING: "Recurring",
};

const HEADERS = {
  // foreignSymbol/foreignCode = the selling-market currency (default VND). currencySymbol/Code = home/base (default JPY).
  Settings: ["currencySymbol","currencyCode","vatRate","monthlyIncomeGoal","monthlyProfitGoal","minLiquidity","fxRateVND","fixedRateVND","foreignSymbol","foreignCode"],
  VATTimeline: ["effectiveFrom", "rate", "note"],
  Income: ["id","date","description","client","category","amountGross","amountVat","amountNet","isVatApplicable","currency","fxRate"],
  Expenses: ["id","date","description","vendor","category","costType","amountGross","amountVat","amountNet","isVatApplicable","linkedPOId"],
  Stock: ["sku","name","category","unit","quantityOnHand","reorderPoint","reorderQuantity","unitCost","unitPrice","lastInboundDate","lastOutboundDate","supplierId"],
  StockMoves: ["moveId", "sku", "date", "qtyDelta", "reason", "refId", "user"],
  PurchaseOrders: ["poId","supplierId","supplierName","orderDate","expectedDate","actualDeliveryDate","currency","fxRate","subtotal","vatAmount","totalGross","status","paymentDueDate","paidDate","notes","paidAmount"],
  POLineItems: ["poId", "sku", "quantity", "unitCost", "vatApplicable"],
  Suppliers: ["supplierId", "name", "contact", "paymentTermsDays", "rating"],
  AuditLog: ["timestamp","user","action","entity","entityId","beforeHash","afterHash","payload"],
  // VND -> JPY currency conversions (treasury). rate = VND per 1 JPY.
  Conversions: ["id","date","vndAmount","rate","jpyReceived","fee","note"],
  // Personal credit cards. Balance is derived from CardTxns (opening + charges - payments).
  CreditCards: ["cardId","name","issuer","creditLimit","statementDay","dueDay","openingBalance"],
  // Card charges (deferred, no cash impact) and payments (cash outflow on date).
  CardTxns: ["id","cardId","date","description","category","amount","type"],
  // Recurring auto-charges (subscriptions, utilities). lastPosted = "YYYY-MM" of the last month auto-posted.
  Recurring: ["id","cardId","name","category","amount","dayOfMonth","active","lastPosted"],
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
      // Migration: add any new columns appended to HEADERS since the sheet was first created
      const currentLastCol = sh.getLastColumn();
      if (currentLastCol < targetHeaders.length) {
        for (let i = currentLastCol; i < targetHeaders.length; i++) {
          sh.getRange(1, i + 1).setValue(targetHeaders[i]).setFontWeight("bold");
        }
      }
    }
  });

  const settings = ss.getSheetByName(TABS.SETTINGS);
  if (settings.getLastRow() === 1) {
    settings.appendRow(["¥", "JPY", 10, 10000000, 4000000, 500000, 0.0058, 185, "₫", "VND"]);
  } else {
    const fillIfBlank = (key, val) => {
      const c = HEADERS.Settings.indexOf(key) + 1;
      const cur = settings.getRange(2, c).getValue();
      if (cur === "" || cur === null || cur === undefined) settings.getRange(2, c).setValue(val);
    };
    fillIfBlank("fxRateVND", 0.0058);
    fillIfBlank("fixedRateVND", 185);   // foreign units per 1 base
    fillIfBlank("foreignSymbol", "₫");
    fillIfBlank("foreignCode", "VND");
  }

  const vat = ss.getSheetByName(TABS.VAT_TIMELINE);
  if (vat.getLastRow() === 1) {
    vat.appendRow(["1997-04-01", 5, "Initial standard rate"]);
    vat.appendRow(["2014-04-01", 8, "Increase to 8%"]);
    vat.appendRow(["2019-10-01", 10, "Current standard rate"]);
  }

  SpreadsheetApp.flush();
  return "Workbook synced with ID: " + id;
}

// =============================================================
// LOW-LEVEL HELPERS
// =============================================================
function _sheet(name) {
  const ss = SpreadsheetApp.openById(_getSsId());
  let sh = ss.getSheetByName(name);
  if (!sh) {
    setupWorkbook();
    sh = ss.getSheetByName(name);
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

function _appendObject(sheetName, obj) {
  const sh = _sheet(sheetName);
  const row = HEADERS[sheetName].map((h) => obj[h] !== undefined ? obj[h] : "");
  sh.appendRow(row);
}

function _updateById(sheetName, idKey, idValue, patch) {
  const sh = _sheet(sheetName);
  const headers = HEADERS[sheetName];
  const colIdx = headers.indexOf(idKey) + 1;
  const last = sh.getLastRow();
  if (last < 2) return false;
  const ids = sh.getRange(2, colIdx, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idValue)) {
      const rowNum = i + 2;
      Object.keys(patch).forEach((k) => {
        const c = headers.indexOf(k) + 1;
        if (c > 0) sh.getRange(rowNum, c).setValue(patch[k]);
      });
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

function _vatRateForDate(dateStr) {
  const cache = CacheService.getScriptCache();
  let timeline = cache.get("vatTimeline");
  if (timeline) { timeline = JSON.parse(timeline); } 
  else {
    timeline = _rows(TABS.VAT_TIMELINE).map((r) => ({
      effectiveFrom: String(r.effectiveFrom),
      rate: Number(r.rate),
    })).sort((a, b) => a.effectiveFrom.localeCompare(b.effectiveFrom));
    cache.put("vatTimeline", JSON.stringify(timeline), 300);
  }
  let rate = 0;
  for (const period of timeline) { if (dateStr >= period.effectiveFrom) rate = period.rate; }
  return rate;
}

function _splitVat(amountGross, isVatApplicable, dateStr) {
  if (!isVatApplicable) return { amountVat: 0, amountNet: amountGross };
  const rate = _vatRateForDate(dateStr);
  const vat = amountGross * (rate / (100 + rate));
  return { amountVat: Math.round(vat), amountNet: Math.round(amountGross - vat) };
}

// =============================================================
// PUBLIC API
// =============================================================

function getBootstrap() {
  const safe = (label, fn, fallback) => {
    try { return fn(); }
    catch (e) { return fallback; }
  };
  // Resolve the database spreadsheet, and self-heal the schema only when a tab is
  // missing (e.g. after adding Recurring/Conversions). getBootstrap runs on every
  // refresh, so we avoid the cost of a full setup unless something is actually absent.
  let ssId;
  try {
    ssId = _getSsId();
    const ss = SpreadsheetApp.openById(ssId);
    const missing = Object.keys(HEADERS).some((t) => !ss.getSheetByName(t));
    if (missing) setupWorkbook();
  } catch (e) {
    return { ok: false, error: 'Cannot open the database spreadsheet (' + (e.message || e) + '). '
      + 'Open your Google Sheet and run "🚀 ProfitIntel → Initialize / Sync Web App", '
      + 'or set the Script Property SPREADSHEET_ID, then reload.' };
  }
  safe('PostDueRecurring', () => postDueRecurring(), null);  // auto-apply subscriptions/utilities
  return {
    ok: true,
    spreadsheetId: ssId,
    serverTime: new Date().toISOString(),
    settings:    safe('Settings',    () => _rows(TABS.SETTINGS)[0] || null, null),
    vatTimeline: safe('VATTimeline', () => _rows(TABS.VAT_TIMELINE),  []),
    income:      safe('Income',      () => _rows(TABS.INCOME),         []),
    expenses:    safe('Expenses',    () => _rows(TABS.EXPENSES),       []),
    stock:       safe('Stock',       () => _rows(TABS.STOCK),          []),
    stockMoves:  safe('StockMoves',  () => _rows(TABS.STOCK_MOVES),    []),
    pos:         safe('PurchaseOrders', () => _rows(TABS.PURCHASE_ORDERS), []),
    poLines:     safe('POLineItems', () => _rows(TABS.PO_LINE_ITEMS),  []),
    suppliers:   safe('Suppliers',   () => _rows(TABS.SUPPLIERS),      []),
    conversions: safe('Conversions', () => _rows(TABS.CONVERSIONS),    []),
    creditCards: safe('CreditCards', () => _rows(TABS.CREDIT_CARDS),   []),
    cardTxns:    safe('CardTxns',    () => _rows(TABS.CARD_TXNS),      []),
    recurring:   safe('Recurring',   () => _rows(TABS.RECURRING),      [])
  };
}

function ping() {
  return { ok: true, spreadsheetId: _getSsId() };
}

function listAuditLog(limit) {
  const all = _rows(TABS.AUDIT_LOG);
  const sorted = all.slice().sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
  return limit ? sorted.slice(0, Number(limit)) : sorted;
}

function getSettings() { return _rows(TABS.SETTINGS)[0] || null; }
function updateSettings(patch) {
  const sh = _sheet(TABS.SETTINGS);
  const headers = HEADERS.Settings;
  if (sh.getLastRow() < 2) sh.appendRow(headers.map((_) => ""));
  Object.keys(patch).forEach((k) => {
    const c = headers.indexOf(k) + 1;
    if (c > 0) sh.getRange(2, c).setValue(patch[k]);
  });
  CacheService.getScriptCache().remove("vatTimeline");
  _audit("UPDATE", "Settings", "singleton", null, patch);
  return getSettings();
}

function listVatTimeline() { return _rows(TABS.VAT_TIMELINE); }
function addVatPeriod(effectiveFrom, rate, note) {
  _appendObject(TABS.VAT_TIMELINE, { effectiveFrom, rate, note: note || "" });
  CacheService.getScriptCache().remove("vatTimeline");
  _audit("CREATE", "VATTimeline", effectiveFrom, null, { effectiveFrom, rate, note });
  return listVatTimeline();
}

function addIncome(tx) {
  const id = tx.id || _uuid();
  const { amountVat, amountNet } = _splitVat(Number(tx.amountGross), tx.isVatApplicable !== false, tx.date);
  const row = { id, date: tx.date, description: tx.description || "", client: tx.client || "", category: tx.category || "Uncategorized", amountGross: Number(tx.amountGross), amountVat, amountNet, isVatApplicable: tx.isVatApplicable !== false, currency: tx.currency || "JPY", fxRate: tx.fxRate || 1 };
  _appendObject(TABS.INCOME, row);
  _audit("CREATE", "Income", id, null, row);
  return row;
}

function addExpense(tx) {
  const id = tx.id || _uuid();
  const { amountVat, amountNet } = _splitVat(Number(tx.amountGross), tx.isVatApplicable !== false, tx.date);
  const row = { id, date: tx.date, description: tx.description || "", vendor: tx.vendor || "", category: tx.category || "Uncategorized", costType: tx.costType || "Variable", amountGross: Number(tx.amountGross), amountVat, amountNet, isVatApplicable: tx.isVatApplicable !== false, linkedPOId: tx.linkedPOId || "" };
  _appendObject(TABS.EXPENSES, row);
  _audit("CREATE", "Expense", id, null, row);
  return row;
}

function deleteTransaction(kind, id) {
  const tabName = kind === "income" ? TABS.INCOME : TABS.EXPENSES;
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
}

function listIncome() { return _rows(TABS.INCOME); }
function listExpenses() { return _rows(TABS.EXPENSES); }
function listStock() { return _rows(TABS.STOCK); }

function upsertStockItem(item) {
  const existing = _rows(TABS.STOCK).find((s) => s.sku === item.sku);
  if (existing) {
    _updateById(TABS.STOCK, "sku", item.sku, item);
    _audit("UPDATE", "Stock", item.sku, existing, item);
  } else {
    _appendObject(TABS.STOCK, item);
    _audit("CREATE", "Stock", item.sku, null, item);
  }
  return item;
}

function adjustStock(sku, qtyDelta, reason, refId) {
  const stock = _rows(TABS.STOCK);
  const item = stock.find((s) => s.sku === sku);
  if (!item) throw new Error("SKU not found: " + sku);
  const newQty = Number(item.quantityOnHand) + Number(qtyDelta);
  const patch = { quantityOnHand: newQty };
  if (qtyDelta > 0) patch.lastInboundDate = new Date().toISOString().slice(0, 10);
  else patch.lastOutboundDate = new Date().toISOString().slice(0, 10);
  _updateById(TABS.STOCK, "sku", sku, patch);
  _appendObject(TABS.STOCK_MOVES, { moveId: _uuid(), sku, date: new Date().toISOString(), qtyDelta, reason: reason || "Manual", refId: refId || "", user: Session.getActiveUser().getEmail() || "unknown" });
  return { sku, quantityOnHand: newQty };
}

const PO_FLOW = { 'Preparing':'Delivery', 'Delivery':'Not Payment', 'Not Payment':'Done', 'Done':null };

function listPurchaseOrders() {
  const headers = _rows(TABS.PURCHASE_ORDERS);
  const lines = _rows(TABS.PO_LINE_ITEMS);
  return headers.map((po) => Object.assign({}, po, { lineItems: lines.filter((l) => l.poId === po.poId) }));
}

function createPO(po) {
  const poId = po.poId || "PO-" + new Date().getFullYear() + "-" + String(Date.now()).slice(-5);
  const orderDate = po.orderDate || new Date().toISOString().slice(0, 10);
  let subtotal = 0, vatAmount = 0;
  (po.lineItems || []).forEach((li) => {
    const lineTotal = Number(li.quantity) * Number(li.unitCost);
    subtotal += lineTotal;
    if (li.vatApplicable) {
      const rate = _vatRateForDate(orderDate);
      vatAmount += lineTotal * (rate / 100);
    }
  });
  const header = { poId, supplierId: po.supplierId || "", supplierName: po.supplierName || "", orderDate, expectedDate: po.expectedDate || "", actualDeliveryDate: "", currency: po.currency || "JPY", fxRate: po.fxRate || 1, subtotal: Math.round(subtotal), vatAmount: Math.round(vatAmount), totalGross: Math.round(subtotal + vatAmount), status: "Preparing", paymentDueDate: po.paymentDueDate || "", paidDate: "", notes: po.notes || "", paidAmount: Math.round(Number(po.depositAmount) || 0) };
  _appendObject(TABS.PURCHASE_ORDERS, header);
  (po.lineItems || []).forEach((li) => _appendObject(TABS.PO_LINE_ITEMS, { poId, sku: li.sku, quantity: li.quantity, unitCost: li.unitCost, vatApplicable: !!li.vatApplicable }));
  _audit("CREATE", "PurchaseOrder", poId, null, header);
  return header;
}

/** Record a (partial) payment against an order. Tracks deposit + balance. */
function recordPOPayment(poId, amount) {
  const pos = _rows(TABS.PURCHASE_ORDERS);
  const po = pos.find((p) => p.poId === poId);
  if (!po) throw new Error("PO not found: " + poId);
  const add = Number(amount) || 0;
  const newPaid = Math.round((Number(po.paidAmount) || 0) + add);
  const patch = { paidAmount: newPaid };
  // Auto-stamp the paid date once fully settled.
  if (newPaid >= Number(po.totalGross) && !po.paidDate) patch.paidDate = new Date().toISOString().slice(0, 10);
  _updateById(TABS.PURCHASE_ORDERS, "poId", poId, patch);
  _audit("UPDATE", "PurchaseOrder", poId, { paidAmount: po.paidAmount }, patch);
  return Object.assign({}, po, patch);
}

function advancePOStatus(poId, newStatus, opts) {
  opts = opts || {};
  const pos = _rows(TABS.PURCHASE_ORDERS);
  const po = pos.find((p) => p.poId === poId);
  if (!po) throw new Error("PO not found: " + poId);
  const patch = { status: newStatus };
  if (newStatus === "Not Payment") {
    patch.actualDeliveryDate = opts.actualDeliveryDate || new Date().toISOString().slice(0, 10);
    const lines = _rows(TABS.PO_LINE_ITEMS).filter((l) => l.poId === poId);
    lines.forEach((l) => adjustStock(l.sku, Number(l.quantity), "PO-Receipt", poId));
    addExpense({ date: patch.actualDeliveryDate, description: "Accrued: " + (po.supplierName || po.supplierId), vendor: po.supplierName, category: "Goods Received Not Invoiced", costType: "Variable", amountGross: Number(po.totalGross), isVatApplicable: true, linkedPOId: poId });
  }
  if (newStatus === "Done") {
    patch.paidDate = opts.paidDate || new Date().toISOString().slice(0, 10);
    addExpense({ date: patch.paidDate, description: "PAID: " + (po.supplierName || po.supplierId), vendor: po.supplierName, category: "Supplier Payment", costType: "Variable", amountGross: Number(po.totalGross), isVatApplicable: true, linkedPOId: poId });
  }
  _updateById(TABS.PURCHASE_ORDERS, "poId", poId, patch);
  _audit("UPDATE", "PurchaseOrder", poId, { status: po.status }, patch);
  return Object.assign({}, po, patch);
}

function listSuppliers() { return _rows(TABS.SUPPLIERS); }
function upsertSupplier(s) {
  const existing = _rows(TABS.SUPPLIERS).find((x) => x.supplierId === s.supplierId);
  if (existing) { _updateById(TABS.SUPPLIERS, "supplierId", s.supplierId, s); _audit("UPDATE", "Supplier", s.supplierId, existing, s); }
  else { _appendObject(TABS.SUPPLIERS, s); _audit("CREATE", "Supplier", s.supplierId, null, s); }
  return s;
}

// =============================================================
// CURRENCY CONVERSIONS (VND -> JPY treasury)
// =============================================================
function listConversions() { return _rows(TABS.CONVERSIONS); }
function addConversion(tx) {
  const id = tx.id || _uuid();
  const vndAmount = Number(tx.vndAmount) || 0;
  const rate = Number(tx.rate) || 0;            // VND per 1 JPY
  const fee = Number(tx.fee) || 0;              // JPY fee
  const jpyReceived = rate > 0 ? Math.round(vndAmount / rate) : 0;
  const row = { id, date: tx.date, vndAmount, rate, jpyReceived, fee, note: tx.note || "" };
  _appendObject(TABS.CONVERSIONS, row);
  _audit("CREATE", "Conversion", id, null, row);
  return row;
}
function saveConversion(tx) { if (tx.id) _deleteRow(TABS.CONVERSIONS, "id", tx.id); return addConversion(tx); }
function deleteConversion(id) { _audit("DELETE", "Conversion", id, { id }, null); return _deleteRow(TABS.CONVERSIONS, "id", id); }

// =============================================================
// CREDIT CARDS (personal cash-flow)
// =============================================================
function listCreditCards() { return _rows(TABS.CREDIT_CARDS); }
function upsertCreditCard(c) {
  const cardId = c.cardId || _uuid();
  const row = { cardId, name: c.name || "Card", issuer: c.issuer || "", creditLimit: Number(c.creditLimit) || 0, statementDay: Number(c.statementDay) || 1, dueDay: Number(c.dueDay) || 10, openingBalance: Number(c.openingBalance) || 0 };
  const existing = _rows(TABS.CREDIT_CARDS).find((x) => x.cardId === cardId);
  if (existing) { _updateById(TABS.CREDIT_CARDS, "cardId", cardId, row); _audit("UPDATE", "CreditCard", cardId, existing, row); }
  else { _appendObject(TABS.CREDIT_CARDS, row); _audit("CREATE", "CreditCard", cardId, null, row); }
  return row;
}
function deleteCreditCard(cardId) {
  _deleteRow(TABS.CREDIT_CARDS, "cardId", cardId);
  // Remove this card's transactions too
  const sh = _sheet(TABS.CARD_TXNS); const last = sh.getLastRow();
  if (last >= 2) { const col = HEADERS.CardTxns.indexOf("cardId") + 1; const ids = sh.getRange(2, col, last - 1, 1).getValues(); for (let i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(cardId)) sh.deleteRow(i + 2); } }
  _audit("DELETE", "CreditCard", cardId, { cardId }, null);
  return true;
}
function listCardTxns() { return _rows(TABS.CARD_TXNS); }
function addCardTxn(tx) {
  const id = tx.id || _uuid();
  const row = { id, cardId: tx.cardId, date: tx.date, description: tx.description || "", category: tx.category || "Other", amount: Number(tx.amount) || 0, type: tx.type === "payment" ? "payment" : "charge" };
  _appendObject(TABS.CARD_TXNS, row);
  _audit("CREATE", "CardTxn", id, null, row);
  return row;
}
function saveCardTxn(tx) { if (tx.id) _deleteRow(TABS.CARD_TXNS, "id", tx.id); return addCardTxn(tx); }
function deleteCardTxn(id) { _audit("DELETE", "CardTxn", id, { id }, null); return _deleteRow(TABS.CARD_TXNS, "id", id); }

// =============================================================
// RECURRING AUTO-CHARGES (subscriptions, utilities)
// =============================================================
function listRecurring() { return _rows(TABS.RECURRING); }
function upsertRecurring(r) {
  const id = r.id || _uuid();
  const row = { id, cardId: r.cardId || "", name: r.name || "Subscription", category: r.category || "Subscription",
    amount: Number(r.amount) || 0, dayOfMonth: Math.min(28, Math.max(1, Number(r.dayOfMonth) || 1)),
    active: r.active === false ? false : true, lastPosted: r.lastPosted || "" };
  const existing = _rows(TABS.RECURRING).find((x) => x.id === id);
  if (existing) { _updateById(TABS.RECURRING, "id", id, row); _audit("UPDATE", "Recurring", id, existing, row); }
  else { _appendObject(TABS.RECURRING, row); _audit("CREATE", "Recurring", id, null, row); }
  return row;
}
function deleteRecurring(id) { _audit("DELETE", "Recurring", id, { id }, null); return _deleteRow(TABS.RECURRING, "id", id); }

/**
 * Post any due recurring charges as card transactions. Idempotent: each rule
 * tracks lastPosted ("YYYY-MM") so a given month is only charged once. Called
 * automatically from getBootstrap, and exposed for a manual "Run now" button.
 */
function postDueRecurring() {
  const rules = _rows(TABS.RECURRING);
  if (!rules.length) return { posted: 0 };
  const now = new Date();
  const curYM = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2);  // local "YYYY-MM"
  const curDay = now.getDate();
  let posted = 0;
  rules.forEach((r) => {
    if (r.active === false || String(r.active).toLowerCase() === "false") return;
    if (!r.cardId || !(Number(r.amount) > 0)) return;
    const day = Math.min(28, Math.max(1, Number(r.dayOfMonth) || 1));
    // Determine months to post: every month strictly after lastPosted, up to current.
    // Current month only counts once its dayOfMonth has arrived.
    const start = r.lastPosted ? _nextYM(String(r.lastPosted).slice(0, 7)) : curYM;
    let ym = start;
    let guard = 0;                       // hard stop: never loop more than ~10 years
    while (ym <= curYM && guard++ < 120) {
      const isCurrent = (ym === curYM);
      if (!isCurrent || curDay >= day) {
        const dateStr = ym + "-" + String(day).padStart(2, "0");
        addCardTxn({ cardId: r.cardId, date: dateStr, type: "charge", amount: Number(r.amount),
          category: r.category || "Subscription", description: "[Auto] " + (r.name || "Recurring") });
        _updateById(TABS.RECURRING, "id", r.id, { lastPosted: ym });
        posted++;
      }
      if (isCurrent) break;
      ym = _nextYM(ym);
    }
  });
  return { posted };
}
// Advance a "YYYY-MM" string by one month using integer math only.
// (Using Date()+toISOString() here is timezone-unsafe: in JST the 1st of a month
//  rolls back to the previous month in UTC, which caused an infinite loop.)
function _nextYM(ym) {
  let y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7)); // m is 1-based
  m += 1; if (m > 12) { m = 1; y += 1; }
  return y + "-" + ("0" + m).slice(-2);
}

function exportSnapshot() {
  const data = { exportedAt: new Date().toISOString(), settings: _rows(TABS.SETTINGS), vatTimeline: _rows(TABS.VAT_TIMELINE), income: _rows(TABS.INCOME), expenses: _rows(TABS.EXPENSES), stock: _rows(TABS.STOCK), stockMoves: _rows(TABS.STOCK_MOVES), pos: _rows(TABS.PURCHASE_ORDERS), poLines: _rows(TABS.PO_LINE_ITEMS), suppliers: _rows(TABS.SUPPLIERS), auditLog: _rows(TABS.AUDIT_LOG) };
  data.signature = _sha256(JSON.stringify(data));
  return data;
}

function seedDemoData() {
  try { resetAllData(); } catch(e) {}
  const today = new Date().toISOString().slice(0,10);
  // Japanese suppliers
  [{supplierId:"SUP-001",name:"Yodobashi Camera",contact:"wholesale@yodobashi.jp",paymentTermsDays:0,rating:"A+"},
   {supplierId:"SUP-002",name:"Mercari Japan",contact:"orders@mercari.jp",paymentTermsDays:0,rating:"A"}].forEach(upsertSupplier);
  // Products: unitCost = ¥ paid in Japan, unitPrice = ₫ charged in Vietnam
  [{sku:"JP-CAM01",name:"Used Mirrorless Camera",category:"Finished Good",unit:"pcs",quantityOnHand:6,reorderPoint:3,reorderQuantity:5,unitCost:42000,unitPrice:11000000,supplierId:"SUP-001",lastInboundDate:today,lastOutboundDate:today},
   {sku:"JP-WAT02",name:"Vintage Seiko Watch",category:"Finished Good",unit:"pcs",quantityOnHand:10,reorderPoint:4,reorderQuantity:8,unitCost:18000,unitPrice:4600000,supplierId:"SUP-002",lastInboundDate:today,lastOutboundDate:today},
   {sku:"JP-SKN03",name:"Skincare Set (Shiseido)",category:"Finished Good",unit:"box",quantityOnHand:30,reorderPoint:10,reorderQuantity:20,unitCost:6500,unitPrice:1750000,supplierId:"SUP-001",lastInboundDate:today,lastOutboundDate:today}].forEach(upsertStockItem);
  // Sales in Vietnam (outbound stock moves drive VND revenue + best sellers)
  adjustStock("JP-CAM01", -3, "Sale", "");
  adjustStock("JP-WAT02", -5, "Sale", "");
  adjustStock("JP-SKN03", -12, "Sale", "");
  // VND -> JPY conversions (fixed selling rate is 185 ₫/¥)
  addConversion({date:today, vndAmount:20000000, rate:178, fee:3000, note:"Converted early — VND strong, good rate"});
  addConversion({date:today, vndAmount:10000000, rate:192, fee:2500, note:"Needed cash — converted at a worse rate"});
  // Personal life
  addIncome({date:today,description:"Salary",client:"Day job",category:"Services",amountGross:280000});
  addExpense({date:today,description:"Rent",vendor:"Landlord",category:"Rent",costType:"Fixed",amountGross:85000});
  // Credit card — buy stock on credit, then a partial payment
  const card = upsertCreditCard({name:"Rakuten Card", issuer:"Rakuten", creditLimit:500000, statementDay:27, dueDay:10, openingBalance:0});
  addCardTxn({cardId:card.cardId, date:today, type:"charge", amount:60000, category:"Goods Purchase", description:"Bought camera stock on credit"});
  addCardTxn({cardId:card.cardId, date:today, type:"charge", amount:23000, category:"Dining", description:"Restaurant"});
  addCardTxn({cardId:card.cardId, date:today, type:"payment", amount:40000, category:"Other", description:"Partial card payment"});
  // Recurring auto-charges (subscriptions / utilities) — auto-posted by postDueRecurring()
  upsertRecurring({cardId:card.cardId, name:"Netflix", category:"Subscription", amount:1980, dayOfMonth:5, active:true});
  upsertRecurring({cardId:card.cardId, name:"Electricity Bill", category:"Utilities", amount:8500, dayOfMonth:25, active:true});
  postDueRecurring();
  // An order with a 30% deposit, awaiting the balance after the customer receives goods
  var demoPO = createPO({supplierId:"SUP-001", supplierName:"Yodobashi Camera", orderDate:today, expectedDate:today,
    paymentDueDate:today, depositAmount: Math.round(5*42000*1.1*0.3),
    lineItems:[{sku:"JP-CAM01", quantity:5, unitCost:42000, vatApplicable:true}]});
  advancePOStatus(demoPO.poId, "Delivery", {});
  advancePOStatus(demoPO.poId, "Not Payment", {});
  return "Demo data seeded.";
}

function saveIncome(tx) { if (tx.id) _deleteRow(TABS.INCOME, "id", tx.id); return addIncome(tx); }
function saveExpense(tx) { if (tx.id) _deleteRow(TABS.EXPENSES, "id", tx.id); return addExpense(tx); }
function deleteIncome(id) { _audit("DELETE", "Income", id, { id }, null); return _deleteRow(TABS.INCOME, "id", id); }
function deleteExpense(id) { _audit("DELETE", "Expense", id, { id }, null); return _deleteRow(TABS.EXPENSES, "id", id); }
function deleteStockItem(sku) { _audit("DELETE", "Stock", sku, { sku }, null); return _deleteRow(TABS.STOCK, "sku", sku); }
function deleteSupplier(supplierId) { _audit("DELETE", "Supplier", supplierId, { supplierId }, null); return _deleteRow(TABS.SUPPLIERS, "supplierId", supplierId); }
function deletePO(poId) { _deleteRow(TABS.PURCHASE_ORDERS, "poId", poId); const sh = _sheet(TABS.PO_LINE_ITEMS); const last = sh.getLastRow(); if (last >= 2) { const ids = sh.getRange(2, 1, last - 1, 1).getValues(); for (let i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(poId)) sh.deleteRow(i + 2); } } _audit("DELETE", "PurchaseOrder", poId, { poId }, null); return true; }

function setStockQty(sku, qty) { const item = _rows(TABS.STOCK).find((s) => s.sku === sku); if (!item) throw new Error("SKU not found: " + sku); const delta = Number(qty) - Number(item.quantityOnHand); if (delta === 0) return { sku, quantityOnHand: Number(qty) }; return adjustStock(sku, delta, "Manual", ""); }
function resetAllData() {
  ["Income","Expenses","Stock","StockMoves","PurchaseOrders","POLineItems","Suppliers","Conversions","CreditCards","CardTxns","Recurring"].forEach((tabName) => {
    try {
      const sh = _sheet(tabName);
      const last = sh.getLastRow();
      const cols = Math.max(1, sh.getLastColumn());
      // Clear the data rows (keep the header). clearContent avoids the
      // "rows out of bounds" error that deleteRows throws when a frozen
      // header row would be the only row left.
      if (last > 1) sh.getRange(2, 1, last - 1, cols).clearContent();
    } catch (e) { /* skip a tab that can't be cleared */ }
  });
  _audit("DELETE", "ALL", "*", null, null);
  return "All transactional data cleared.";
}

function _deleteRow(sheetName, idKey, idValue) { const sh = _sheet(sheetName); const headers = HEADERS[sheetName]; const colIdx = headers.indexOf(idKey) + 1; const last = sh.getLastRow(); if (last < 2) return false; const ids = sh.getRange(2, colIdx, last - 1, 1).getValues(); for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(idValue)) { sh.deleteRow(i + 2); return true; } } return false; }
