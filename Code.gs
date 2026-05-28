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
};

const HEADERS = {
  Settings: ["currencySymbol","currencyCode","vatRate","monthlyIncomeGoal","monthlyProfitGoal","minLiquidity"],
  VATTimeline: ["effectiveFrom", "rate", "note"],
  Income: ["id","date","description","client","category","amountGross","amountVat","amountNet","isVatApplicable","currency","fxRate"],
  Expenses: ["id","date","description","vendor","category","costType","amountGross","amountVat","amountNet","isVatApplicable","linkedPOId"],
  Stock: ["sku","name","category","unit","quantityOnHand","reorderPoint","reorderQuantity","unitCost","unitPrice","lastInboundDate","lastOutboundDate","supplierId"],
  StockMoves: ["moveId", "sku", "date", "qtyDelta", "reason", "refId", "user"],
  PurchaseOrders: ["poId","supplierId","supplierName","orderDate","expectedDate","actualDeliveryDate","currency","fxRate","subtotal","vatAmount","totalGross","status","paymentDueDate","paidDate","notes"],
  POLineItems: ["poId", "sku", "quantity", "unitCost", "vatApplicable"],
  Suppliers: ["supplierId", "name", "contact", "paymentTermsDays", "rating"],
  AuditLog: ["timestamp","user","action","entity","entityId","beforeHash","afterHash","payload"],
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = ss.getId();
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", id);
  
  Object.keys(HEADERS).forEach((tabName) => {
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[tabName]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[tabName].length).setFontWeight("bold");
    }
  });

  const settings = ss.getSheetByName(TABS.SETTINGS);
  if (settings.getLastRow() === 1) {
    settings.appendRow(["¥", "JPY", 10, 10000000, 4000000, 500000]);
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
  return {
    ok: true,
    spreadsheetId: _getSsId(),
    serverTime: new Date().toISOString(),
    settings:    safe('Settings',    () => _rows(TABS.SETTINGS)[0] || null, null),
    vatTimeline: safe('VATTimeline', () => _rows(TABS.VAT_TIMELINE),  []),
    income:      safe('Income',      () => _rows(TABS.INCOME),         []),
    expenses:    safe('Expenses',    () => _rows(TABS.EXPENSES),       []),
    stock:       safe('Stock',       () => _rows(TABS.STOCK),          []),
    stockMoves:  safe('StockMoves',  () => _rows(TABS.STOCK_MOVES),    []),
    pos:         safe('PurchaseOrders', () => _rows(TABS.PURCHASE_ORDERS), []),
    poLines:     safe('POLineItems', () => _rows(TABS.PO_LINE_ITEMS),  []),
    suppliers:   safe('Suppliers',   () => _rows(TABS.SUPPLIERS),      [])
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
  const header = { poId, supplierId: po.supplierId || "", supplierName: po.supplierName || "", orderDate, expectedDate: po.expectedDate || "", actualDeliveryDate: "", currency: po.currency || "JPY", fxRate: po.fxRate || 1, subtotal: Math.round(subtotal), vatAmount: Math.round(vatAmount), totalGross: Math.round(subtotal + vatAmount), status: "Preparing", paymentDueDate: po.paymentDueDate || "", paidDate: "", notes: po.notes || "" };
  _appendObject(TABS.PURCHASE_ORDERS, header);
  (po.lineItems || []).forEach((li) => _appendObject(TABS.PO_LINE_ITEMS, { poId, sku: li.sku, quantity: li.quantity, unitCost: li.unitCost, vatApplicable: !!li.vatApplicable }));
  _audit("CREATE", "PurchaseOrder", poId, null, header);
  return header;
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

function exportSnapshot() {
  const data = { exportedAt: new Date().toISOString(), settings: _rows(TABS.SETTINGS), vatTimeline: _rows(TABS.VAT_TIMELINE), income: _rows(TABS.INCOME), expenses: _rows(TABS.EXPENSES), stock: _rows(TABS.STOCK), stockMoves: _rows(TABS.STOCK_MOVES), pos: _rows(TABS.PURCHASE_ORDERS), poLines: _rows(TABS.PO_LINE_ITEMS), suppliers: _rows(TABS.SUPPLIERS), auditLog: _rows(TABS.AUDIT_LOG) };
  data.signature = _sha256(JSON.stringify(data));
  return data;
}

function seedDemoData() {
  try { resetAllData(); } catch(e) {}
  [{supplierId:"SUP-001",name:"Nippon Steel",contact:"sato@nipponsteel.jp",paymentTermsDays:30,rating:"A+"},{supplierId:"SUP-002",name:"Kobe Steel",contact:"tanaka@kobelco.jp",paymentTermsDays:30,rating:"A"},{supplierId:"SUP-003",name:"Mitsui Chemicals",contact:"orders@mitsui.jp",paymentTermsDays:45,rating:"B"}].forEach(upsertSupplier);
  [{sku:"SKU-1001",name:"Steel Beam 3m",category:"Raw Material",unit:"pcs",quantityOnHand:120,reorderPoint:50,reorderQuantity:100,unitCost:18500,unitPrice:24000,supplierId:"SUP-001",lastInboundDate:"2026-05-15",lastOutboundDate:"2026-05-26"},{sku:"SKU-2010",name:"Assembled Frame X",category:"Finished Good",unit:"pcs",quantityOnHand:25,reorderPoint:15,reorderQuantity:30,unitCost:124000,unitPrice:185000,supplierId:"SUP-001",lastInboundDate:"2026-05-18",lastOutboundDate:"2026-05-27"}].forEach(upsertStockItem);
  addIncome({date:"2026-05-28",description:"Q2 Licensing Revenue",client:"Acme Corp",category:"Licensing",amountGross:3300000});
  addExpense({date:"2026-05-27",description:"AWS Cloud Hosting",vendor:"Amazon Web Services",category:"Software",costType:"Variable",amountGross:412500});
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
function resetAllData() { ["Income","Expenses","Stock","StockMoves","PurchaseOrders","POLineItems","Suppliers"].forEach((tabName) => { const sh = _sheet(tabName); const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }); _audit("DELETE", "ALL", "*", null, null); return "All transactional data cleared."; }

function _deleteRow(sheetName, idKey, idValue) { const sh = _sheet(sheetName); const headers = HEADERS[sheetName]; const colIdx = headers.indexOf(idKey) + 1; const last = sh.getLastRow(); if (last < 2) return false; const ids = sh.getRange(2, colIdx, last - 1, 1).getValues(); for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(idValue)) { sh.deleteRow(i + 2); return true; } } return false; }
