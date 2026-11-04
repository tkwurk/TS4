/**
 * Stock history — cumulative totals through yesterday in Stock_History;
 * stock report at runtime = history snapshot + today-only deltas.
 * Bootstrap: one-shot read (Purchases + Consumption + Orders/BOM) through yesterday.
 * Nightly: roll forward one day when snapshot is exactly one day behind; else one-shot catch-up.
 */
var STOCK_HISTORY_HEADERS = [
  'AsOfDate',
  'Outlet',
  'ProductCode',
  'TotalPurchased',
  'TotalConsumed',
  'TotalSold',
  'UpdatedAt'
];

var STOCK_HISTORY_SOURCE_NIGHTLY = 'NIGHTLY';
var STOCK_TOTALS_CACHE_TTL_SEC = 300;

function _todayBusinessDateStr_() {
  return formatDateForComparison(new Date());
}

function _nextBusinessDateStr_(dateStr) {
  if (!dateStr) return null;
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() + 1);
  return formatDateForComparison(d);
}

function _prevBusinessDateStr_(dateStr) {
  if (!dateStr) return null;
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() - 1);
  return formatDateForComparison(d);
}

function _roundStockQty_(n) {
  return Math.round((Number(n) || 0) * 1000) / 1000;
}

function _normalizeStockCode_(code) {
  return typeof normalizeProductCodeKey_ === 'function'
    ? normalizeProductCodeKey_(code)
    : String(code || '').trim();
}

function ensureStockHistorySheet_() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STOCK_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STOCK_HISTORY);
    sheet.appendRow(STOCK_HISTORY_HEADERS);
    Logger.log('Created ' + SHEETS.STOCK_HISTORY + ' sheet');
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(STOCK_HISTORY_HEADERS);
  }
  return sheet;
}

/**
 * Sum purchase qty by product code through throughDateStr (inclusive).
 */
function _readPurchaseQtyByCodeThroughDate_(outlet, throughDateStr) {
  var purchaseQty = {};
  var throughKey = String(throughDateStr || '').trim();
  try {
    var purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (!purchasesSheet || purchasesSheet.getLastRow() < 2) return purchaseQty;
    var purchaseData = purchasesSheet.getDataRange().getValues();
    var phHeaders = purchaseData.shift();
    var phDateIdx = phHeaders.findIndex(function(h) { return h && /date|timestamp/i.test(h.toString()); });
    var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
    // Historic stock uses received quantity (RQUANTITY, col T), not ordered Quantity (col E).
    var phRQtyIdx = phHeaders.findIndex(function(h) { return h && /rquantity/i.test(h.toString()); });
    var phQtyIdx = phRQtyIdx >= 0
      ? phRQtyIdx
      : phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
    var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
    purchaseData.forEach(function(row) {
      if (phOutletIdx !== -1 && outlet) {
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
        } else if (row[phOutletIdx] !== outlet) return;
      }
      if (phDateIdx >= 0 && row[phDateIdx] && throughKey) {
        var rawDate = row[phDateIdx];
        var parsed = typeof parsePurchaseDate === 'function'
          ? parsePurchaseDate(rawDate)
          : (rawDate instanceof Date ? rawDate : new Date(rawDate));
        if (!parsed || isNaN(parsed.getTime())) return;
        var dateStr = formatDateForComparison(parsed);
        if (!dateStr || dateStr > throughKey) return;
      }
      var code = phCodeIdx >= 0 ? (row[phCodeIdx] || '').toString().trim() : '';
      var qty = Number(row[phQtyIdx]) || 0;
      if (!code || qty === 0) return;
      var nk = _normalizeStockCode_(code);
      purchaseQty[nk] = (purchaseQty[nk] || 0) + qty;
    });
  } catch (e) {
    Logger.log('_readPurchaseQtyByCodeThroughDate_ error: ' + e.toString());
  }
  return purchaseQty;
}

/**
 * Sum recorded inventory consumption by product code through throughDateStr (inclusive).
 */
function _readRecordedConsumptionQtyByCodeThroughDate_(outlet, throughDateStr) {
  var consumptionQty = {};
  var throughKey = String(throughDateStr || '').trim();
  try {
    var consumptionSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!consumptionSheet || consumptionSheet.getLastRow() < 2) return consumptionQty;
    var consData = consumptionSheet.getDataRange().getValues();
    var ch = consData[0];
    var cDateIdx = ch.findIndex(function(h) { return h && /timestamp|date/i.test(h.toString()); });
    var cCodeIdx = ch.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
    var cQtyIdx = ch.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
    var cOutletIdx = ch.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
    for (var ci = 1; ci < consData.length; ci++) {
      var crow = consData[ci];
      if (cOutletIdx !== -1 && outlet) {
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(crow[cOutletIdx], outlet)) continue;
        } else if (crow[cOutletIdx] !== outlet) continue;
      }
      if (cDateIdx >= 0 && crow[cDateIdx] && throughKey) {
        var ts = crow[cDateIdx] instanceof Date ? crow[cDateIdx] : new Date(crow[cDateIdx]);
        if (isNaN(ts.getTime())) continue;
        var dateStr = formatDateForComparison(ts);
        if (!dateStr || dateStr > throughKey) continue;
      }
      var cCode = cCodeIdx >= 0 ? (crow[cCodeIdx] || '').toString().trim() : '';
      var cQty = cQtyIdx >= 0 ? (Number(crow[cQtyIdx]) || 0) : 0;
      if (!cCode || cQty === 0) continue;
      var nk = _normalizeStockCode_(cCode);
      consumptionQty[nk] = (consumptionQty[nk] || 0) + cQty;
    }
  } catch (e) {
    Logger.log('_readRecordedConsumptionQtyByCodeThroughDate_ error: ' + e.toString());
  }
  return consumptionQty;
}

/**
 * One read each of Purchases, Consumption, Orders → cumulative totals through throughDateStr.
 */
function buildStockHistoryByCodeThroughDate_(outlet, throughDateStr, options) {
  options = options || {};
  var purchased = _readPurchaseQtyByCodeThroughDate_(outlet, throughDateStr);
  var consumed = _readRecordedConsumptionQtyByCodeThroughDate_(outlet, throughDateStr);
  var sold = {};

  if (typeof computeBomConsumptionThroughDateForOutlet_ === 'function') {
    var bomResult = computeBomConsumptionThroughDateForOutlet_(outlet, throughDateStr, options);
    sold = bomResult.consumptionMap || {};
  }

  var byCode = {};
  var codes = {};
  Object.keys(purchased).forEach(function(k) { codes[k] = true; });
  Object.keys(consumed).forEach(function(k) { codes[k] = true; });
  Object.keys(sold).forEach(function(k) { codes[k] = true; });

  Object.keys(codes).forEach(function(code) {
    byCode[code] = {
      totalPurchased: _roundStockQty_(purchased[code] || 0),
      totalConsumed: _roundStockQty_(consumed[code] || 0),
      totalSold: _roundStockQty_(sold[code] || 0)
    };
  });

  return byCode;
}

/**
 * One read of Purchases → { 'yyyy-MM-dd': { productCode: qty } } per outlet.
 */
function buildPurchaseDateIndexForOutlet_(outlet) {
  var index = {};
  try {
    var purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (!purchasesSheet || purchasesSheet.getLastRow() < 2) return index;
    var purchaseData = purchasesSheet.getDataRange().getValues();
    var phHeaders = purchaseData.shift();
    var phDateIdx = phHeaders.findIndex(function(h) { return h && /date|timestamp/i.test(h.toString()); });
    var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
    // Historic stock uses received quantity (RQUANTITY, col T), not ordered Quantity (col E).
    var phRQtyIdx = phHeaders.findIndex(function(h) { return h && /rquantity/i.test(h.toString()); });
    var phQtyIdx = phRQtyIdx >= 0
      ? phRQtyIdx
      : phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
    var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
    purchaseData.forEach(function(row) {
      if (phOutletIdx !== -1 && outlet) {
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
        } else if (row[phOutletIdx] !== outlet) return;
      }
      if (phDateIdx < 0 || !row[phDateIdx]) return;
      var rawDate = row[phDateIdx];
      var parsed = typeof parsePurchaseDate === 'function'
        ? parsePurchaseDate(rawDate)
        : (rawDate instanceof Date ? rawDate : new Date(rawDate));
      if (!parsed || isNaN(parsed.getTime())) return;
      var dateStr = formatDateForComparison(parsed);
      if (!dateStr) return;
      var code = phCodeIdx >= 0 ? (row[phCodeIdx] || '').toString().trim() : '';
      var qty = Number(row[phQtyIdx]) || 0;
      if (!code || qty === 0) return;
      var nk = _normalizeStockCode_(code);
      if (!index[dateStr]) index[dateStr] = {};
      index[dateStr][nk] = (index[dateStr][nk] || 0) + qty;
    });
  } catch (e) {
    Logger.log('buildPurchaseDateIndexForOutlet_ error: ' + e.toString());
  }
  return index;
}

/**
 * Purchases + recorded consumption + BOM sold for one business date (today-only path).
 */
function computeStockDeltasForDate_(outlet, businessDateStr, options) {
  options = options || {};
  var purchased = {};
  if (options.purchaseDateIndex) {
    purchased = options.purchaseDateIndex[businessDateStr] || {};
  } else {
    purchased = _readPurchaseQtyByCodeForDate_(outlet, businessDateStr, options.purchaseDateIndex);
  }

  var consumed = {};
  if (typeof readInventoryConsumptionForDate_ === 'function') {
    var rawConsumed = readInventoryConsumptionForDate_(businessDateStr, outlet).consumptionMap || {};
    Object.keys(rawConsumed).forEach(function(k) {
      var nk = _normalizeStockCode_(k);
      if (!nk) return;
      consumed[nk] = (consumed[nk] || 0) + (Number(rawConsumed[k]) || 0);
    });
  }

  var sold = {};
  if (typeof computeCombinedConsumptionForOutlet_ === 'function') {
    var bomRecipes = options.bomRecipes || (typeof getCachedBomRecipes_ === 'function'
      ? getCachedBomRecipes_()
      : (typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : { byUnitProduct: {} }));
    var usageMatrix = options.usageMatrix || (typeof getCachedBomUsageMatrix_ === 'function'
      ? getCachedBomUsageMatrix_()
      : (typeof loadActiveBomUsageMatrix_ === 'function' ? loadActiveBomUsageMatrix_() : {}));
    var productNames = options.productNames || (typeof getCachedProductNamesByCode_ === 'function'
      ? getCachedProductNamesByCode_()
      : (typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {}));
    var combined = computeCombinedConsumptionForOutlet_(
      businessDateStr,
      outlet,
      bomRecipes,
      usageMatrix,
      productNames,
      new Date(),
      {
        productIndex: options.productIndex,
        preferBottomUp: options.preferBottomUp,
        batchSize: options.batchSize
      }
    );
    var rawSold = combined.consumptionMap || {};
    Object.keys(rawSold).forEach(function(k) {
      var nk = _normalizeStockCode_(k);
      if (!nk) return;
      sold[nk] = (sold[nk] || 0) + (Number(rawSold[k]) || 0);
    });
  }

  return {
    businessDate: businessDateStr,
    purchased: purchased,
    consumed: consumed,
    sold: sold
  };
}

function _readPurchaseQtyByCodeForDate_(outlet, businessDateStr, purchaseDateIndex) {
  if (purchaseDateIndex) {
    return purchaseDateIndex[businessDateStr] || {};
  }
  var purchaseQty = {};
  try {
    var purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (!purchasesSheet || purchasesSheet.getLastRow() < 2) return purchaseQty;
    var purchaseData = purchasesSheet.getDataRange().getValues();
    var phHeaders = purchaseData.shift();
    var phDateIdx = phHeaders.findIndex(function(h) { return h && /date|timestamp/i.test(h.toString()); });
    var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
    // Historic stock uses received quantity (RQUANTITY, col T), not ordered Quantity (col E).
    var phRQtyIdx = phHeaders.findIndex(function(h) { return h && /rquantity/i.test(h.toString()); });
    var phQtyIdx = phRQtyIdx >= 0
      ? phRQtyIdx
      : phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
    var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
    purchaseData.forEach(function(row) {
      if (phOutletIdx !== -1 && outlet) {
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
        } else if (row[phOutletIdx] !== outlet) return;
      }
      if (phDateIdx < 0 || !row[phDateIdx]) return;
      var rawDate = row[phDateIdx];
      var parsed = typeof parsePurchaseDate === 'function'
        ? parsePurchaseDate(rawDate)
        : (rawDate instanceof Date ? rawDate : new Date(rawDate));
      if (!parsed || isNaN(parsed.getTime())) return;
      if (formatDateForComparison(parsed) !== businessDateStr) return;
      var code = phCodeIdx >= 0 ? (row[phCodeIdx] || '').toString().trim() : '';
      var qty = Number(row[phQtyIdx]) || 0;
      if (!code || qty === 0) return;
      var nk = _normalizeStockCode_(code);
      purchaseQty[nk] = (purchaseQty[nk] || 0) + qty;
    });
  } catch (e) {
    Logger.log('_readPurchaseQtyByCodeForDate_ error: ' + e.toString());
  }
  return purchaseQty;
}

function readStockHistorySnapshot_(outlet) {
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.STOCK_HISTORY);
  if (!sheet || sheet.getLastRow() < 2) {
    return { asOfDate: null, byCode: {} };
  }

  var outletKey = String(outlet || '').trim();
  var lastRow = sheet.getLastRow();
  var numCols = Math.max(sheet.getLastColumn(), STOCK_HISTORY_HEADERS.length);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var latestDate = null;

  for (var i = data.length - 1; i >= 0; i--) {
    var rowOutlet = String(data[i][1] || '').trim();
    if (outletKey && rowOutlet !== outletKey) continue;
    var asOf = String(data[i][0] || '').trim();
    if (!asOf) continue;
    if (!latestDate || asOf > latestDate) latestDate = asOf;
  }

  if (!latestDate) {
    return { asOfDate: null, byCode: {} };
  }

  var byCode = {};
  data.forEach(function(row) {
    if (String(row[1] || '').trim() !== outletKey) return;
    if (String(row[0] || '').trim() !== latestDate) return;
    var code = _normalizeStockCode_(row[2]);
    if (!code) return;
    byCode[code] = {
      totalPurchased: _roundStockQty_(row[3]),
      totalConsumed: _roundStockQty_(row[4]),
      totalSold: _roundStockQty_(row[5])
    };
  });

  return { asOfDate: latestDate, byCode: byCode };
}

function replaceStockHistorySnapshot_(outlet, asOfDate, byCode) {
  ensureStockHistorySheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.STOCK_HISTORY);
  if (!sheet) return 0;

  var outletKey = String(outlet || '').trim();
  var lastRow = sheet.getLastRow();
  var numCols = STOCK_HISTORY_HEADERS.length;
  var kept = [];

  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    kept = data.filter(function(row) {
      return String(row[1] || '').trim() !== outletKey;
    });
  }

  var genStr = formatTimestampForSheet(new Date());
  var newRows = [];
  Object.keys(byCode || {}).forEach(function(code) {
    var item = byCode[code];
    if (!item) return;
    newRows.push([
      asOfDate,
      outletKey,
      code,
      _roundStockQty_(item.totalPurchased),
      _roundStockQty_(item.totalConsumed),
      _roundStockQty_(item.totalSold),
      genStr
    ]);
  });

  var allRows = kept.concat(newRows);
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
  }
  if (allRows.length) {
    sheet.getRange(2, 1, allRows.length, numCols).setValues(allRows);
  } else if (lastRow >= 2) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return newRows.length;
}

function _applyDeltaToSnapshotByCode_(byCode, delta) {
  byCode = byCode || {};
  var codes = {};
  Object.keys(byCode).forEach(function(k) { codes[k] = true; });
  Object.keys(delta.purchased || {}).forEach(function(k) { codes[k] = true; });
  Object.keys(delta.consumed || {}).forEach(function(k) { codes[k] = true; });
  Object.keys(delta.sold || {}).forEach(function(k) { codes[k] = true; });

  Object.keys(codes).forEach(function(code) {
    var cur = byCode[code] || { totalPurchased: 0, totalConsumed: 0, totalSold: 0 };
    byCode[code] = {
      totalPurchased: _roundStockQty_(cur.totalPurchased + (Number(delta.purchased[code]) || 0)),
      totalConsumed: _roundStockQty_(cur.totalConsumed + (Number(delta.consumed[code]) || 0)),
      totalSold: _roundStockQty_(cur.totalSold + (Number(delta.sold[code]) || 0))
    };
  });
  return byCode;
}

/**
 * Update Stock_History through throughDateStr.
 * Bootstrap / catch-up: one-shot totals from beginning through target date.
 * Nightly (exactly one day behind): add single-day delta only.
 */
function advanceStockHistoryForOutlet_(outlet, throughDateStr, options) {
  options = options || {};
  throughDateStr = throughDateStr || (typeof getYesterdayBusinessDateStr_ === 'function'
    ? getYesterdayBusinessDateStr_()
    : _todayBusinessDateStr_());

  var snapshot = readStockHistorySnapshot_(outlet);
  if (snapshot.asOfDate && snapshot.asOfDate >= throughDateStr) {
    return {
      success: true,
      outlet: outlet,
      asOfDate: snapshot.asOfDate,
      rows: 0,
      daysApplied: 0,
      skipped: true,
      reason: 'already_current',
      complete: true,
      mode: 'skip'
    };
  }

  var nextDay = snapshot.asOfDate ? _nextBusinessDateStr_(snapshot.asOfDate) : null;
  var byCode;
  var mode;

  if (snapshot.asOfDate && nextDay === throughDateStr) {
    var purchaseDateIndex = options.purchaseDateIndex || buildPurchaseDateIndexForOutlet_(outlet);
    options.purchaseDateIndex = purchaseDateIndex;
    var delta = computeStockDeltasForDate_(outlet, throughDateStr, options);
    byCode = _applyDeltaToSnapshotByCode_(snapshot.byCode || {}, delta);
    mode = 'delta';
  } else {
    byCode = buildStockHistoryByCodeThroughDate_(outlet, throughDateStr, options);
    mode = 'one_shot';
  }

  var rowCount = replaceStockHistorySnapshot_(outlet, throughDateStr, byCode);

  return {
    success: true,
    outlet: outlet,
    asOfDate: throughDateStr,
    rows: rowCount,
    daysApplied: mode === 'delta' ? 1 : 0,
    skipped: false,
    complete: true,
    mode: mode
  };
}

/**
 * Ensure Stock_History for an outlet is current through yesterday. Called on
 * navigation to Daily Closing / Stock Report so the snapshot auto-refreshes
 * without waiting for the nightly trigger. Idempotent, and guarded so the
 * (possibly heavy) catch-up runs at most once per short window per outlet.
 */
function ensureStockHistoryCurrentForOutlet_(outlet, options) {
  options = options || {};
  var outletKey = String(outlet || '').trim();
  if (!outletKey) return { success: false, skipped: true, reason: 'no_outlet' };

  var yesterdayStr = typeof getYesterdayBusinessDateStr_ === 'function'
    ? getYesterdayBusinessDateStr_()
    : _todayBusinessDateStr_();

  var snapshot = readStockHistorySnapshot_(outletKey);
  if (snapshot.asOfDate && snapshot.asOfDate >= yesterdayStr) {
    return { success: true, skipped: true, reason: 'already_current', asOfDate: snapshot.asOfDate };
  }

  var cache = CacheService.getScriptCache();
  var guardKey = 'stock_hist_autorefresh_' + outletKey + '_' + yesterdayStr;
  if (!options.force) {
    try {
      if (cache.get(guardKey)) {
        return { success: true, skipped: true, reason: 'recently_attempted' };
      }
    } catch (e) { /* ignore cache read errors */ }
  }
  try { cache.put(guardKey, '1', 120); } catch (e2) { /* ignore */ }

  try {
    var bomRecipes = typeof getCachedBomRecipes_ === 'function'
      ? getCachedBomRecipes_()
      : (typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : null);
    var usageMatrix = typeof getCachedBomUsageMatrix_ === 'function'
      ? getCachedBomUsageMatrix_()
      : (typeof loadActiveBomUsageMatrix_ === 'function' ? loadActiveBomUsageMatrix_() : null);
    var productNames = typeof getCachedProductNamesByCode_ === 'function'
      ? getCachedProductNamesByCode_()
      : (typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {});
    return advanceStockHistoryForOutlet_(outletKey, yesterdayStr, {
      bomRecipes: bomRecipes,
      usageMatrix: usageMatrix,
      productNames: productNames
    });
  } catch (err) {
    Logger.log('ensureStockHistoryCurrentForOutlet_ error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Fast fallback: cached one-read each of Purchases + Inventory_Consumption (no order scan).
 */
function getStockTotalsDirectSheet_(outlet, options) {
  options = options || {};
  var outletKey = String(outlet || '').trim();
  var cacheKey = 'stock_direct_v1_' + outletKey;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      var parsed = JSON.parse(cached);
      if (parsed && parsed.purchased) {
        parsed.stockSource = 'DirectSheet';
        parsed.needsBootstrap = true;
        parsed.historyAsOfDate = null;
        return parsed;
      }
    }
  } catch (e) { /* ignore cache parse errors */ }

  var purchased = typeof _readPurchaseQtyByCode_ === 'function'
    ? _readPurchaseQtyByCode_(outlet) : {};
  var consumed = typeof _readRecordedConsumptionQtyByCode_ === 'function'
    ? _readRecordedConsumptionQtyByCode_(outlet) : {};
  var result = {
    purchased: purchased,
    consumed: consumed,
    sold: {},
    stockSource: 'DirectSheet',
    needsBootstrap: true,
    historyAsOfDate: null
  };
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), STOCK_TOTALS_CACHE_TTL_SEC);
  } catch (e2) { /* cache payload too large — skip */ }
  return result;
}

function getCachedStockTotalsForOutlet_(outlet, options) {
  options = options || {};
  // Do NOT cache the combined totals — today's deltas (sales, purchases,
  // consumption) must appear immediately. The history portion (through
  // yesterday) is already persisted in Stock_History, so live computation only
  // scans today's activity and stays fast.
  return getStockTotalsForOutlet_(outlet, options);
}

/**
 * Fast stock totals: history (if current through yesterday) + today only.
 * Never runs bootstrap inline — avoids execution timeout on stock report.
 */
function getStockTotalsForOutlet_(outlet, options) {
  options = options || {};
  var todayStr = _todayBusinessDateStr_();
  var yesterdayStr = typeof getYesterdayBusinessDateStr_ === 'function'
    ? getYesterdayBusinessDateStr_()
    : todayStr;

  var snapshot = readStockHistorySnapshot_(outlet);

  if (!snapshot.asOfDate || snapshot.asOfDate < yesterdayStr) {
    var direct = getStockTotalsDirectSheet_(outlet, options);
    // The direct/fallback path carries no sold figures; purchased/consumed are
    // already lifetime, so merge today's sold to reflect the current day while
    // Stock_History catches up.
    var fallbackDelta = computeStockDeltasForDate_(outlet, todayStr, options);
    var fallbackSold = {};
    Object.keys(fallbackDelta.sold || {}).forEach(function(k) {
      fallbackSold[k] = _roundStockQty_(Number(fallbackDelta.sold[k]) || 0);
    });
    direct.sold = fallbackSold;
    direct.todayDate = todayStr;
    direct.historyIncomplete = !!snapshot.asOfDate;
    direct.historyAsOfDate = snapshot.asOfDate;
    return direct;
  }

  var todayDelta = computeStockDeltasForDate_(outlet, todayStr, options);
  var purchased = {};
  var consumed = {};
  var sold = {};
  var codes = {};

  Object.keys(snapshot.byCode || {}).forEach(function(k) { codes[k] = true; });
  Object.keys(todayDelta.purchased || {}).forEach(function(k) { codes[k] = true; });
  Object.keys(todayDelta.consumed || {}).forEach(function(k) { codes[k] = true; });
  Object.keys(todayDelta.sold || {}).forEach(function(k) { codes[k] = true; });

  Object.keys(codes).forEach(function(code) {
    var hist = (snapshot.byCode && snapshot.byCode[code]) || {};
    purchased[code] = _roundStockQty_((hist.totalPurchased || 0) + (Number(todayDelta.purchased[code]) || 0));
    consumed[code] = _roundStockQty_((hist.totalConsumed || 0) + (Number(todayDelta.consumed[code]) || 0));
    sold[code] = _roundStockQty_((hist.totalSold || 0) + (Number(todayDelta.sold[code]) || 0));
  });

  return {
    purchased: purchased,
    consumed: consumed,
    sold: sold,
    stockSource: 'History+Today',
    historyAsOfDate: snapshot.asOfDate,
    todayDate: todayStr,
    needsBootstrap: false
  };
}

function runNightlyStockHistoryJob(businessDateStr, options) {
  options = options || {};
  ensureStockHistorySheet_();

  var businessDate = businessDateStr || (typeof getYesterdayBusinessDateStr_ === 'function'
    ? getYesterdayBusinessDateStr_()
    : _todayBusinessDateStr_());
  var force = !!options.force;
  var onlyOutlet = options.outlet ? String(options.outlet).trim() : '';

  var outlets = typeof getDistinctOutletsForConsumption_ === 'function'
    ? getDistinctOutletsForConsumption_()
    : [];
  if (onlyOutlet) outlets = [onlyOutlet];

  var bomRecipes = typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : null;
  var usageMatrix = typeof loadActiveBomUsageMatrix_ === 'function' ? loadActiveBomUsageMatrix_() : null;
  var productNames = typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {};
  var sharedOpts = {
    bomRecipes: bomRecipes,
    usageMatrix: usageMatrix,
    productNames: productNames
  };

  var results = [];
  var totalRows = 0;

  outlets.forEach(function(outlet) {
    var snapshot = readStockHistorySnapshot_(outlet);
    if (!force && snapshot.asOfDate && snapshot.asOfDate >= businessDate) {
      results.push({
        outlet: outlet,
        skipped: true,
        reason: 'already_summarized',
        rows: 0,
        asOfDate: snapshot.asOfDate,
        complete: true,
        mode: 'skip'
      });
      return;
    }

    var result = advanceStockHistoryForOutlet_(outlet, businessDate, sharedOpts);
    totalRows += result.rows || 0;
    results.push(result);
  });

  var props = PropertiesService.getScriptProperties();
  props.setProperty('STOCK_HISTORY_LAST_RUN', formatTimestampForSheet(new Date()));
  props.setProperty('STOCK_HISTORY_LAST_DATE', businessDate);

  return {
    success: true,
    date: businessDate,
    totalRows: totalRows,
    outletsProcessed: results.length,
    results: results,
    complete: true,
    message: 'Updated stock history through ' + businessDate + ' (' + totalRows + ' row(s))'
  };
}

function scheduleNightlyStockHistoryTrigger(hour) {
  if (hour === undefined || hour === null) hour = 1;
  hour = Math.max(0, Math.min(23, parseInt(hour, 10) || 0));
  removeNightlyStockHistoryTrigger();
  ScriptApp.newTrigger('runNightlyStockHistoryJob')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
  return {
    success: true,
    message: 'Nightly stock history trigger set for ' + (hour < 10 ? '0' : '') + hour + ':00'
  };
}

function removeNightlyStockHistoryTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runNightlyStockHistoryJob') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return { success: true, removed: removed };
}

function getNightlyStockHistoryTriggerInfo() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = null;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runNightlyStockHistoryJob') {
      found = {
        id: trigger.getUniqueId(),
        type: trigger.getEventType().toString(),
        source: trigger.getTriggerSource().toString()
      };
    }
  });
  var props = PropertiesService.getScriptProperties();
  return {
    triggerActive: !!found,
    trigger: found,
    lastRun: props.getProperty('STOCK_HISTORY_LAST_RUN') || null,
    lastBusinessDate: props.getProperty('STOCK_HISTORY_LAST_DATE') || null,
    sheetName: SHEETS.STOCK_HISTORY
  };
}

function runNightlyStockHistoryJobSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  var businessDate = params && params.businessDate ? String(params.businessDate).trim() : '';
  var force = !!(params && params.force);
  var outlet = params && params.outlet ? String(params.outlet).trim() : '';
  return runNightlyStockHistoryJob(businessDate || undefined, { force: force, outlet: outlet || undefined });
}

function rebuildStockHistoryFullSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  var outlet = params && params.outlet ? String(params.outlet).trim() : '';
  if (!outlet) {
    return { success: false, error: 'outlet required' };
  }
  replaceStockHistorySnapshot_(outlet, null, {});
  var yesterdayStr = typeof getYesterdayBusinessDateStr_ === 'function'
    ? getYesterdayBusinessDateStr_()
    : _todayBusinessDateStr_();
  var bomRecipes = typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : null;
  var usageMatrix = typeof loadActiveBomUsageMatrix_ === 'function' ? loadActiveBomUsageMatrix_() : null;
  var productNames = typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {};
  return advanceStockHistoryForOutlet_(outlet, yesterdayStr, {
    bomRecipes: bomRecipes,
    usageMatrix: usageMatrix,
    productNames: productNames,
    force: true
  });
}

function scheduleNightlyStockHistoryTriggerSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  var hour = params && params.hour !== undefined ? params.hour : 1;
  return scheduleNightlyStockHistoryTrigger(hour);
}

function removeNightlyStockHistoryTriggerSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  return removeNightlyStockHistoryTrigger();
}

function getNightlyStockHistoryTriggerInfoSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  return getNightlyStockHistoryTriggerInfo();
}
