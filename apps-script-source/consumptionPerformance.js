// ========== CONSUMPTION PERFORMANCE (cache, order scan, daily summary) ==========

var CONSUMPTION_CACHE_TTL_SEC = 300;
var CONSUMPTION_CACHE_VERSION = 'v1';
var CONSUMPTION_ORDER_BATCH_SIZE = 500;

var _consumptionRequestCache_ = null;

function getConsumptionRequestCache_() {
  if (!_consumptionRequestCache_) _consumptionRequestCache_ = {};
  return _consumptionRequestCache_;
}

function clearConsumptionRequestCache_() {
  _consumptionRequestCache_ = null;
}

function consumptionCacheKey_(suffix) {
  var ver = typeof SCHEMA_VERSION !== 'undefined' ? SCHEMA_VERSION : 'v1';
  return 'cons_' + CONSUMPTION_CACHE_VERSION + '_' + ver + '_' + suffix;
}

function getJsonFromScriptCache_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function putJsonInScriptCache_(key, value, ttlSec) {
  try {
    var serialized = JSON.stringify(value);
    if (serialized.length >= 95000) return false;
    CacheService.getScriptCache().put(key, serialized, ttlSec || CONSUMPTION_CACHE_TTL_SEC);
    return true;
  } catch (e) {
    return false;
  }
}

function invalidateConsumptionCache_() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(consumptionCacheKey_('bom'));
    cache.remove(consumptionCacheKey_('usage'));
    cache.remove(consumptionCacheKey_('product_index'));
    cache.remove(consumptionCacheKey_('product_names'));
    cache.remove(consumptionCacheKey_('buy_products'));
  } catch (e) {
    Logger.log('invalidateConsumptionCache_ failed: ' + e.message);
  }
  clearConsumptionRequestCache_();
}

function getCachedBomRecipes_() {
  var req = getConsumptionRequestCache_();
  if (req.bomRecipes) return req.bomRecipes;

  var cached = getJsonFromScriptCache_(consumptionCacheKey_('bom'));
  if (cached && cached.byUnitProduct) {
    req.bomRecipes = cached;
    return cached;
  }

  var loaded = typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : {
    byUnitProduct: {}, byUnitLookup: {}, lines: [], totalLines: 0, skipped: []
  };
  putJsonInScriptCache_(consumptionCacheKey_('bom'), loaded, CONSUMPTION_CACHE_TTL_SEC);
  req.bomRecipes = loaded;
  return loaded;
}

function getCachedBomUsageMatrix_() {
  var req = getConsumptionRequestCache_();
  if (req.usageMatrix) return req.usageMatrix;

  var cached = getJsonFromScriptCache_(consumptionCacheKey_('usage'));
  if (cached && cached.byBaseAndFlavor) {
    req.usageMatrix = cached;
    return cached;
  }

  var loaded = typeof loadActiveBomUsageMatrix_ === 'function'
    ? loadActiveBomUsageMatrix_()
    : { byBaseAndFlavor: {}, baseCodes: {}, flavorCodes: {}, lines: [], skipped: [] };
  putJsonInScriptCache_(consumptionCacheKey_('usage'), loaded, CONSUMPTION_CACHE_TTL_SEC);
  req.usageMatrix = loaded;
  return loaded;
}

function getCachedProductCodeIndex_() {
  var req = getConsumptionRequestCache_();
  if (req.productIndex) return req.productIndex;

  var cached = getJsonFromScriptCache_(consumptionCacheKey_('product_index'));
  if (cached && cached.byCode) {
    req.productIndex = cached;
    return cached;
  }

  var loaded = typeof loadProductCodeIndex_ === 'function'
    ? loadProductCodeIndex_()
    : { byCode: {}, byGenerated: {}, byName: {} };
  putJsonInScriptCache_(consumptionCacheKey_('product_index'), loaded, CONSUMPTION_CACHE_TTL_SEC);
  req.productIndex = loaded;
  return loaded;
}

function getCachedProductNamesByCode_() {
  var req = getConsumptionRequestCache_();
  if (req.productNames) return req.productNames;

  var cached = getJsonFromScriptCache_(consumptionCacheKey_('product_names'));
  if (cached) {
    req.productNames = cached;
    return cached;
  }

  var loaded = typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {};
  putJsonInScriptCache_(consumptionCacheKey_('product_names'), loaded, CONSUMPTION_CACHE_TTL_SEC);
  req.productNames = loaded;
  return loaded;
}

function getCachedBuyTypeProducts_() {
  var req = getConsumptionRequestCache_();
  if (req.buyProducts) return req.buyProducts;

  var cached = getJsonFromScriptCache_(consumptionCacheKey_('buy_products'));
  if (cached && Array.isArray(cached)) {
    req.buyProducts = cached;
    return cached;
  }

  var loaded = typeof getBuyTypeProducts === 'function' ? getBuyTypeProducts() : [];
  putJsonInScriptCache_(consumptionCacheKey_('buy_products'), loaded, CONSUMPTION_CACHE_TTL_SEC);
  req.buyProducts = loaded;
  return loaded;
}

function getYesterdayBusinessDateStr_() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return formatDateForComparison(d);
}

/**
 * Build daily-closing prefill items for one outlet + business date.
 * Uses cached BOM/products and a single order scan.
 */
function buildConsumptionPrefillItems_(businessDateStr, outlet, options) {
  options = options || {};
  var bomRecipes = options.bomRecipes || getCachedBomRecipes_();
  var usageMatrix = options.usageMatrix || getCachedBomUsageMatrix_();
  var productNames = options.productNames || getCachedProductNamesByCode_();
  var buyProducts = options.buyProducts || getCachedBuyTypeProducts_();

  var buyByCode = {};
  buyProducts.forEach(function(p) {
    var k = normalizeProductCodeKey_(p.productCode);
    if (k) buyByCode[k] = p;
  });

  var preferBottomUp = options.preferBottomUp;
  if (preferBottomUp === undefined && typeof shouldPreferBottomUpOrderScan_ === 'function') {
    preferBottomUp = shouldPreferBottomUpOrderScan_(businessDateStr);
  }

  var combined = typeof computeCombinedConsumptionForOutlet_ === 'function'
    ? computeCombinedConsumptionForOutlet_(
        businessDateStr, outlet, bomRecipes, usageMatrix, productNames, new Date(),
        {
          productIndex: getCachedProductCodeIndex_(),
          preferBottomUp: preferBottomUp,
          batchSize: options.batchSize
        }
      )
    : { consumptionMap: {}, soldMap: {}, details: [] };

  if (typeof buildConsumptionPrefillLinesFromDetails_ === 'function') {
    return buildConsumptionPrefillLinesFromDetails_(
      combined.details || [],
      buyByCode,
      productNames
    );
  }

  return [];
}

/**
 * Read Inventory_Consumption rows for one business date + outlet (bottom-up scan).
 */
function readInventoryConsumptionForDate_(businessDateStr, outlet, options) {
  options = options || {};
  var batchSize = CONSUMPTION_ORDER_BATCH_SIZE || 500;
  var firstBatch = Repo.readRowsFromEnd('INVENTORY_CONSUMPTION', 0, batchSize);
  if (!firstBatch.headers.length || !firstBatch.totalRows) {
    return { consumptionMap: {}, rowsScanned: 0 };
  }

  var headerRow = firstBatch.headers;
  var timestampIdx = headerRow.findIndex(function(h) { return h && /timestamp|date/i.test(h.toString()); });
  var productCodeIdx = headerRow.findIndex(function(h) { return h && /product.*code|code/i.test(h.toString()); });
  var quantityIdx = headerRow.findIndex(function(h) { return h && /quantity/i.test(h.toString()); });
  var outletIdx = headerRow.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  if (timestampIdx < 0 || productCodeIdx < 0 || quantityIdx < 0) {
    throw new Error('Inventory_Consumption is missing a date, product code, or quantity column.');
  }
  if (options.strictOutlet && outletIdx < 0) throw new Error('Inventory_Consumption is missing an outlet column.');

  var consumptionMap = {};
  var rowsScanned = 0;
  var outletKey = String(outlet || '').trim();
  var offset = 0;

  while (offset < firstBatch.totalRows) {
    var batchInfo = offset === 0
      ? firstBatch
      : Repo.readRowsFromEnd('INVENTORY_CONSUMPTION', offset, batchSize, firstBatch.totalRows);
    var batch = batchInfo.rows.map(function (row) {
      return headerRow.map(function (header) { return row[header]; });
    });
    for (var i = batch.length - 1; i >= 0; i--) {
      rowsScanned++;
      var row = batch[i];
      if (outletKey && outletIdx >= 0) {
        if (options.strictOutlet) {
          if (canonicalOutletKey_(row[outletIdx]) !== canonicalOutletKey_(outletKey)) continue;
        } else if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(row[outletIdx], outletKey)) continue;
        } else if (String(row[outletIdx] || '').trim() !== outletKey) {
          continue;
        }
      }
      if (timestampIdx < 0 || !row[timestampIdx]) continue;
      var ts = parseSheetDate_(row[timestampIdx]);
      if (!ts) continue;
      var dateStr = typeof getAppDateKey_ === 'function'
        ? getAppDateKey_(ts)
        : formatDateForComparison(ts);
      if (dateStr < businessDateStr) return { consumptionMap: consumptionMap, rowsScanned: rowsScanned };
      if (dateStr !== businessDateStr) continue;

      var productCode = productCodeIdx >= 0 ? (row[productCodeIdx] || '').toString().trim() : '';
      var quantity = quantityIdx >= 0 ? (Number(row[quantityIdx]) || 0) : 0;
      if (!productCode || quantity <= 0) continue;
      consumptionMap[productCode] = (consumptionMap[productCode] || 0) + quantity;
    }
    offset += batch.length;
  }

  return { consumptionMap: consumptionMap, rowsScanned: rowsScanned };
}
