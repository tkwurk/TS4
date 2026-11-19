/**
 * Stock history — cumulative totals through yesterday in Stock_History;
 * the live report adds only current-day deltas from bottom-up scans.
 * Bootstrap: one-shot read (Purchases + Consumption + Orders/BOM) through yesterday.
 * Midnight job: roll forward one day when the snapshot is exactly one day behind;
 * otherwise perform a one-shot catch-up.
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

function _stockBusinessDateKey_(date) {
  return typeof getAppDateKey_ === 'function'
    ? getAppDateKey_(date)
    : formatDateForComparison(date);
}

function _stockSnapshotDateKey_(value) {
  if (!value) return '';
  if (value instanceof Date) return _stockBusinessDateKey_(value);
  var dateKey = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
    throw new Error('Invalid Stock_History AsOfDate: ' + dateKey);
  }
  return dateKey;
}

function _todayBusinessDateStr_() {
  return _stockBusinessDateKey_(new Date());
}

function _nextBusinessDateStr_(dateStr) {
  if (!dateStr) return null;
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() + 1);
  return _stockBusinessDateKey_(d);
}

function _prevBusinessDateStr_(dateStr) {
  if (!dateStr) return null;
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() - 1);
  return _stockBusinessDateKey_(d);
}

function _roundStockQty_(n) {
  return Math.round((Number(n) || 0) * 1000) / 1000;
}

function _normalizeStockCode_(code) {
  return typeof normalizeProductCodeKey_ === 'function'
    ? normalizeProductCodeKey_(code)
    : String(code || '').trim();
}

function _stockRepoValues_(collection) {
  var table = Repo.readTable(collection);
  if (!table.headers.length) return [];
  return [table.headers].concat(table.rows.map(function (row) {
    return table.headers.map(function (header) { return row[header]; });
  }));
}

function ensureStockHistorySheet_() {
  Repo.ensureTable('STOCK_HISTORY', STOCK_HISTORY_HEADERS);
  return true;
}

/**
 * Sum purchase qty by product code through throughDateStr (inclusive).
 */
function _readPurchaseQtyByCodeThroughDate_(outlet, throughDateStr) {
  var purchaseQty = {};
  var throughKey = String(throughDateStr || '').trim();
  var purchaseData = _stockRepoValues_('PURCHASES');
  if (purchaseData.length < 2) return purchaseQty;
  var phHeaders = purchaseData[0];
  var phDateIdx = phHeaders.findIndex(function(h) { return h && /date|timestamp/i.test(h.toString()); });
  var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
  // Stock uses received quantity when available, not ordered quantity.
  var phRQtyIdx = phHeaders.findIndex(function(h) { return h && /rquantity/i.test(h.toString()); });
  var phQtyIdx = phRQtyIdx >= 0
    ? phRQtyIdx
    : phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  if (phDateIdx < 0 || phCodeIdx < 0 || phQtyIdx < 0) {
    throw new Error('Purchases is missing a date, product code, or quantity column.');
  }
  purchaseData.slice(1).forEach(function(row) {
    if (phOutletIdx !== -1 && outlet) {
      if (typeof outletRowMatchesUser_ === 'function') {
        if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
      } else if (row[phOutletIdx] !== outlet) return;
    }
    if (row[phDateIdx] && throughKey) {
      var rawDate = row[phDateIdx];
      var parsed = typeof parsePurchaseDate === 'function'
        ? parsePurchaseDate(rawDate)
        : (rawDate instanceof Date ? rawDate : new Date(rawDate));
      if (!parsed || isNaN(parsed.getTime())) return;
      var dateStr = _stockBusinessDateKey_(parsed);
      if (!dateStr || dateStr > throughKey) return;
    }
    var code = (row[phCodeIdx] || '').toString().trim();
    var qty = Number(row[phQtyIdx]) || 0;
    if (!code || qty === 0) return;
    var nk = _normalizeStockCode_(code);
    purchaseQty[nk] = (purchaseQty[nk] || 0) + qty;
  });
  return purchaseQty;
}

/**
 * Sum recorded inventory consumption by product code through throughDateStr (inclusive).
 */
function _readRecordedConsumptionQtyByCodeThroughDate_(outlet, throughDateStr) {
  var consumptionQty = {};
  var throughKey = String(throughDateStr || '').trim();
  var consData = _stockRepoValues_('INVENTORY_CONSUMPTION');
  if (consData.length < 2) return consumptionQty;
  var ch = consData[0];
  var cDateIdx = ch.findIndex(function(h) { return h && /timestamp|date/i.test(h.toString()); });
  var cCodeIdx = ch.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var cQtyIdx = ch.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var cOutletIdx = ch.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  if (cDateIdx < 0 || cCodeIdx < 0 || cQtyIdx < 0) {
    throw new Error('Inventory_Consumption is missing a date, product code, or quantity column.');
  }
  for (var ci = 1; ci < consData.length; ci++) {
    var crow = consData[ci];
    if (cOutletIdx !== -1 && outlet) {
      if (typeof outletRowMatchesUser_ === 'function') {
        if (!outletRowMatchesUser_(crow[cOutletIdx], outlet)) continue;
      } else if (crow[cOutletIdx] !== outlet) continue;
    }
    if (crow[cDateIdx] && throughKey) {
      var ts = crow[cDateIdx] instanceof Date ? crow[cDateIdx] : new Date(crow[cDateIdx]);
      if (isNaN(ts.getTime())) continue;
      var dateStr = _stockBusinessDateKey_(ts);
      if (!dateStr || dateStr > throughKey) continue;
    }
    var cCode = (crow[cCodeIdx] || '').toString().trim();
    var cQty = Number(crow[cQtyIdx]) || 0;
    if (!cCode || cQty === 0) continue;
    var nk = _normalizeStockCode_(cCode);
    consumptionQty[nk] = (consumptionQty[nk] || 0) + cQty;
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
  var purchaseData = _stockRepoValues_('PURCHASES');
  if (purchaseData.length < 2) return index;
  var phHeaders = purchaseData[0];
  var phDateIdx = phHeaders.findIndex(function(h) { return h && /date|timestamp/i.test(h.toString()); });
  var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
  var phRQtyIdx = phHeaders.findIndex(function(h) { return h && /rquantity/i.test(h.toString()); });
  var phQtyIdx = phRQtyIdx >= 0
    ? phRQtyIdx
    : phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  if (phDateIdx < 0 || phCodeIdx < 0 || phQtyIdx < 0) {
    throw new Error('Purchases is missing a date, product code, or quantity column.');
  }
  purchaseData.slice(1).forEach(function(row) {
    if (phOutletIdx !== -1 && outlet) {
      if (typeof outletRowMatchesUser_ === 'function') {
        if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
      } else if (row[phOutletIdx] !== outlet) return;
    }
    if (!row[phDateIdx]) return;
    var rawDate = row[phDateIdx];
    var parsed = typeof parsePurchaseDate === 'function'
      ? parsePurchaseDate(rawDate)
      : (rawDate instanceof Date ? rawDate : new Date(rawDate));
    if (!parsed || isNaN(parsed.getTime())) return;
    var dateStr = _stockBusinessDateKey_(parsed);
    if (!dateStr) return;
    var code = (row[phCodeIdx] || '').toString().trim();
    var qty = Number(row[phQtyIdx]) || 0;
    if (!code || qty === 0) return;
    var nk = _normalizeStockCode_(code);
    if (!index[dateStr]) index[dateStr] = {};
    index[dateStr][nk] = (index[dateStr][nk] || 0) + qty;
  });
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
  var purchaseData = _stockRepoValues_('PURCHASES');
  if (purchaseData.length < 2) return purchaseQty;
  var phHeaders = purchaseData[0];
  var phDateIdx = phHeaders.findIndex(function(h) { return h && /date|timestamp/i.test(h.toString()); });
  var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
  var phRQtyIdx = phHeaders.findIndex(function(h) { return h && /rquantity/i.test(h.toString()); });
  var phQtyIdx = phRQtyIdx >= 0
    ? phRQtyIdx
    : phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  if (phDateIdx < 0 || phCodeIdx < 0 || phQtyIdx < 0) {
    throw new Error('Purchases is missing a date, product code, or quantity column.');
  }
  purchaseData.slice(1).forEach(function(row) {
    if (phOutletIdx !== -1 && outlet) {
      if (typeof outletRowMatchesUser_ === 'function') {
        if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
      } else if (row[phOutletIdx] !== outlet) return;
    }
    if (!row[phDateIdx]) return;
    var rawDate = row[phDateIdx];
    var parsed = typeof parsePurchaseDate === 'function'
      ? parsePurchaseDate(rawDate)
      : (rawDate instanceof Date ? rawDate : new Date(rawDate));
    if (!parsed || isNaN(parsed.getTime())) return;
    if (_stockBusinessDateKey_(parsed) !== businessDateStr) return;
    var code = (row[phCodeIdx] || '').toString().trim();
    var qty = Number(row[phQtyIdx]) || 0;
    if (!code || qty === 0) return;
    var nk = _normalizeStockCode_(code);
    purchaseQty[nk] = (purchaseQty[nk] || 0) + qty;
  });
  return purchaseQty;
}

function readStockHistorySnapshot_(outlet) {
  var data = _stockRepoValues_('STOCK_HISTORY');
  if (data.length < 2) {
    return { asOfDate: null, byCode: {} };
  }

  var outletKey = String(outlet || '').trim();
  var latestDate = null;

  for (var i = data.length - 1; i >= 1; i--) {
    var rowOutlet = String(data[i][1] || '').trim();
    if (outletKey && rowOutlet !== outletKey) continue;
    var asOf = _stockSnapshotDateKey_(data[i][0]);
    if (!asOf) continue;
    if (!latestDate || asOf > latestDate) latestDate = asOf;
  }

  if (!latestDate) {
    return { asOfDate: null, byCode: {} };
  }

  var byCode = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[1] || '').trim() !== outletKey) continue;
    if (_stockSnapshotDateKey_(row[0]) !== latestDate) continue;
    var code = _normalizeStockCode_(row[2]);
    if (!code) continue;
    byCode[code] = {
      totalPurchased: _roundStockQty_(row[3]),
      totalConsumed: _roundStockQty_(row[4]),
      totalSold: _roundStockQty_(row[5])
    };
  }

  return { asOfDate: latestDate, byCode: byCode };
}

function replaceStockHistorySnapshot_(outlet, asOfDate, byCode) {
  ensureStockHistorySheet_();
  var outletKey = String(outlet || '').trim();
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

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var existing = _stockRepoValues_('STOCK_HISTORY');
    var kept = existing.slice(1).filter(function(row) {
      return String(row[1] || '').trim() !== outletKey;
    });
    var allRows = kept.concat(newRows);
    var objects = allRows.map(function (row) {
      var item = {};
      STOCK_HISTORY_HEADERS.forEach(function (header, index) {
        item[header] = row[index];
      });
      return item;
    });
    Repo.replaceRows('STOCK_HISTORY', objects);
  } finally {
    lock.releaseLock();
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
    if (options.requireCompleteHistory) {
      return {
        purchased: {},
        consumed: {},
        sold: {},
        stockSource: 'History+Today',
        historyAsOfDate: snapshot.asOfDate,
        todayDate: todayStr,
        historyIncomplete: true,
        needsBootstrap: true
      };
    }
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
    direct.historyIncomplete = true;
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

  if (businessDateStr && typeof businessDateStr === 'object') {
    businessDateStr = undefined;
  }
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
  if (hour === undefined || hour === null) hour = 0;
  hour = Math.max(0, Math.min(23, parseInt(hour, 10) || 0));
  removeNightlyStockHistoryTrigger();
  ScriptApp.newTrigger('runNightlyStockHistoryJob')
    .timeBased()
    .atHour(hour)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  PropertiesService.getScriptProperties().setProperty('STOCK_HISTORY_SCHEDULE_HOUR', String(hour));
  return {
    success: true,
    message: 'Nightly stock history trigger set near ' + (hour < 10 ? '0' : '') + hour + ':00'
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
  PropertiesService.getScriptProperties().deleteProperty('STOCK_HISTORY_SCHEDULE_HOUR');
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
  var scheduledHour = props.getProperty('STOCK_HISTORY_SCHEDULE_HOUR');
  var historyData = _stockRepoValues_('STOCK_HISTORY');
  var latestByOutlet = {};
  for (var i = 1; i < historyData.length; i++) {
    var outlet = String(historyData[i][1] || '').trim();
    var asOfDate = _stockSnapshotDateKey_(historyData[i][0]);
    if (!outlet || !asOfDate) continue;
    if (!latestByOutlet[outlet] || asOfDate > latestByOutlet[outlet]) {
      latestByOutlet[outlet] = asOfDate;
    }
  }
  var historyThrough = null;
  Object.keys(latestByOutlet).forEach(function (outlet) {
    var asOfDate = latestByOutlet[outlet];
    if (!historyThrough || asOfDate < historyThrough) historyThrough = asOfDate;
  });
  return {
    triggerActive: !!found,
    trigger: found,
    scheduleHour: scheduledHour === null ? null : Number(scheduledHour),
    timeZone: Session.getScriptTimeZone(),
    lastRun: props.getProperty('STOCK_HISTORY_LAST_RUN') || null,
    lastBusinessDate: historyThrough,
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
  var hour = params && params.hour !== undefined ? params.hour : 0;
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
