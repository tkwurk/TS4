// ========== BOM / RECIPE MANAGEMENT ==========
// Maps sell products (Unit Product Code) to ingredient buy products (Recipe Item Code).

var BOM_SCHEDULER_USERNAME = 'BOM_SCHEDULER';

var BOM_HEADERS = [
  'BOM Id',
  'BOM Item Id',
  'Recipe Item Code',
  'Item Quantity',
  'Unit Product Code',
  'Status'
];

/**
 * Ensure BOM sheet exists with expected headers.
 */
function ensureBomSheet_() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.BOM);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.BOM);
    sheet.appendRow(BOM_HEADERS);
    Logger.log('Created BOM sheet');
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(BOM_HEADERS);
    return sheet;
  }
  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), BOM_HEADERS.length)).getValues()[0];
  var normalized = existing.map(function(h) { return String(h || '').trim().toLowerCase(); });
  var expected = BOM_HEADERS.map(function(h) { return h.toLowerCase(); });
  var mismatch = expected.some(function(h, i) { return normalized[i] !== h; });
  if (mismatch) {
    Logger.log('BOM sheet header mismatch. Expected: ' + BOM_HEADERS.join(', ') + ' Got: ' + existing.join(', '));
  }
  return sheet;
}

function isBomRowActive_(status) {
  var s = (status || '').toString().trim().toLowerCase();
  if (!s) return true;
  return s === 'active' || s === 'aktif' || s === 'yes' || s === 'y' || s === '1' || s === 'true' || s === 'enabled';
}

/** Case-insensitive product code key for matching Orders ↔ BOM ↔ Products. */
function normalizeProductCodeKey_(code) {
  return (code || '').toString().trim().toUpperCase().replace(/\s+/g, '_');
}

function _readRepoValues_(collection) {
  var table = Repo.readTable(collection);
  if (!table.headers.length) return [];
  return [table.headers].concat(table.rows.map(function (row) {
    return table.headers.map(function (header) { return row[header]; });
  }));
}

/**
 * Index Products sheet codes, generated codes (category_name), and names.
 */
function loadProductCodeIndex_() {
  var byCode = {};
  var byGenerated = {};
  var byName = {};
  var data = _readRepoValues_('PRODUCTS');
  if (data.length < 2) {
    return { byCode: byCode, byGenerated: byGenerated, byName: byName };
  }

  var headers = data[0];
  var nameIdx = headers.findIndex(function(h) { return h && /^name$/i.test(h.toString()); });
  var catIdx = headers.findIndex(function(h) { return h && /category/i.test(h.toString()); });
  var codeIdx = headers.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var activeIdx = headers.findIndex(function(h) { return h && /^active$/i.test(h.toString()); });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(row[activeIdx])) continue;
    var name = nameIdx >= 0 ? (row[nameIdx] || '').toString().trim() : '';
    var category = catIdx >= 0 ? (row[catIdx] || '').toString().trim() : '';
    var rawCode = codeIdx >= 0 ? (row[codeIdx] || '').toString().trim() : '';
    var generated = typeof generateProductCode === 'function'
      ? generateProductCode({ name: name, category: category, productCode: rawCode })
      : rawCode;
    var codeKey = normalizeProductCodeKey_(rawCode || generated);
    var genKey = normalizeProductCodeKey_(generated);
    if (!codeKey && !genKey) continue;

    var entry = {
      code: codeKey,
      generated: genKey,
      name: name,
      rawCode: rawCode
    };
    if (codeKey) byCode[codeKey] = entry;
    if (genKey) byGenerated[genKey] = entry;
    var nameKey = normalizeProductCodeKey_(name);
    if (nameKey) byName[nameKey] = entry;
  }
  return { byCode: byCode, byGenerated: byGenerated, byName: byName };
}

/**
 * Expand BOM index so sell codes from Orders match even when formatting differs.
 */
function buildBomUnitLookup_(byUnitProduct, productIndex) {
  var lookup = {};
  var index = productIndex || (typeof getCachedProductCodeIndex_ === 'function'
    ? getCachedProductCodeIndex_()
    : loadProductCodeIndex_());

  function register(alias, lines) {
    var k = normalizeProductCodeKey_(alias);
    if (!k || !lines || !lines.length) return;
    if (!lookup[k]) lookup[k] = [];
    lines.forEach(function(line) {
      var exists = lookup[k].some(function(ex) {
        return ex.unitProductCode === line.unitProductCode &&
          ex.recipeItemCode === line.recipeItemCode &&
          ex.bomItemId === line.bomItemId;
      });
      if (!exists) lookup[k].push(line);
    });
  }

  Object.keys(byUnitProduct || {}).forEach(function(unitCode) {
    var lines = byUnitProduct[unitCode];
    register(unitCode, lines);
    var nk = normalizeProductCodeKey_(unitCode);
    var prod = index.byCode[nk] || index.byGenerated[nk];
    if (prod) {
      register(prod.code, lines);
      register(prod.generated, lines);
      register(prod.rawCode, lines);
    }
  });
  return lookup;
}

function resolveBomLinesForSold_(soldCode, bomLookup) {
  if (!bomLookup) return null;
  var k = normalizeProductCodeKey_(soldCode);
  if (bomLookup[k]) return bomLookup[k];

  var index = loadProductCodeIndex_();
  var prod = index.byCode[k] || index.byGenerated[k] || index.byName[k];
  if (!prod) return null;
  if (prod.code && bomLookup[prod.code]) return bomLookup[prod.code];
  if (prod.generated && bomLookup[prod.generated]) return bomLookup[prod.generated];
  return null;
}

function resolveSellToParentForSold_(soldCode, sellToParent) {
  if (!sellToParent) return null;
  var k = normalizeProductCodeKey_(soldCode);
  if (sellToParent[k]) return sellToParent[k];
  if (sellToParent[soldCode]) return sellToParent[soldCode];

  var index = loadProductCodeIndex_();
  var prod = index.byCode[k] || index.byGenerated[k] || index.byName[k];
  if (!prod) return null;
  if (prod.code && sellToParent[prod.code]) return sellToParent[prod.code];
  if (prod.generated && sellToParent[prod.generated]) return sellToParent[prod.generated];
  return null;
}

function resolveSoldCodeFromOrderRow_(row, codeIdx, nameIdx, productIndex) {
  var soldCode = codeIdx >= 0 ? (row[codeIdx] || '').toString().trim() : '';
  if (soldCode) return normalizeProductCodeKey_(soldCode);

  if (nameIdx >= 0) {
    var name = (row[nameIdx] || '').toString().trim();
    var nameKey = normalizeProductCodeKey_(name);
    if (nameKey && productIndex.byName[nameKey]) {
      return productIndex.byName[nameKey].code || productIndex.byName[nameKey].generated;
    }
    if (name && typeof generateProductCode === 'function') {
      return normalizeProductCodeKey_(generateProductCode({ name: name }));
    }
  }
  return '';
}

/**
 * Parse Item Quantity: whole number, decimal, or fraction text (e.g. "1/9").
 * In Sheets you can also use a formula (=1/9) — getValues() returns the decimal.
 */
function parseBomItemQuantity_(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number' && !isNaN(raw) && raw > 0) return raw;

  var s = String(raw).trim();
  if (!s) return 0;

  var asNum = Number(s);
  if (!isNaN(asNum) && asNum > 0) return asNum;

  var frac = s.match(/^(\d+(?:[.,]\d+)?)\s*\/\s*(\d+(?:[.,]\d+)?)$/);
  if (frac) {
    var numer = parseFloat(String(frac[1]).replace(',', '.'));
    var denom = parseFloat(String(frac[2]).replace(',', '.'));
    if (denom > 0 && numer > 0) return numer / denom;
  }

  // Locale decimal comma e.g. "0,05"
  if (/^\d+,\d+$/.test(s)) {
    var commaNum = parseFloat(s.replace(',', '.'));
    if (!isNaN(commaNum) && commaNum > 0) return commaNum;
  }
  return 0;
}

/**
 * Load active BOM lines grouped by unit (sell) product code.
 * @returns {{ byUnitProduct: Object.<string, Array>, lines: Array, totalLines: number }}
 */
function loadActiveBomRecipes_() {
  var byUnitProduct = {};
  var lines = [];
  var skipped = [];
  var data = _readRepoValues_('BOM');
  if (data.length < 2) {
    return { byUnitProduct: byUnitProduct, byUnitLookup: {}, lines: lines, totalLines: 0, skipped: skipped };
  }

  var headers = data[0];
  var bomIdIdx = headers.findIndex(function(h) { return h && /bom\s*id/i.test(h.toString()); });
  var bomItemIdx = headers.findIndex(function(h) { return h && /bom\s*item\s*id/i.test(h.toString()); });
  var recipeIdx = headers.findIndex(function(h) { return h && /recipe\s*item\s*code/i.test(h.toString()); });
  var qtyIdx = headers.findIndex(function(h) {
    return h && /item\s*quantity|item\s*qty/i.test(h.toString());
  });
  if (qtyIdx === -1) {
    qtyIdx = headers.findIndex(function(h) { return h && /^quantity$/i.test(h.toString()); });
  }
  if (qtyIdx === -1 && typeof COLUMNS !== 'undefined' && COLUMNS.BOM && COLUMNS.BOM.ITEM_QUANTITY) {
    qtyIdx = COLUMNS.BOM.ITEM_QUANTITY.col;
  }
  var unitIdx = headers.findIndex(function(h) { return h && /unit\s*product\s*code|sell\s*product\s*code/i.test(h.toString()); });
  var statusIdx = headers.findIndex(function(h) { return h && /^status$/i.test(h.toString()); });

  if (recipeIdx === -1 || unitIdx === -1 || qtyIdx === -1) {
    Logger.log('BOM sheet missing required columns');
    return { byUnitProduct: byUnitProduct, byUnitLookup: {}, lines: lines, totalLines: 0, skipped: skipped };
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var unitCode = (row[unitIdx] || '').toString().trim();
    var recipeCode = (row[recipeIdx] || '').toString().trim();
    var qtyRaw = row[qtyIdx];
    var itemQty = parseBomItemQuantity_(qtyRaw);
    var rowEmpty = !unitCode && !recipeCode && (qtyRaw === '' || qtyRaw === null || qtyRaw === undefined);
    if (rowEmpty) continue;

    if (statusIdx >= 0 && !isBomRowActive_(row[statusIdx])) {
      skipped.push({
        sheetRow: i + 1,
        unitProductCode: unitCode,
        recipeItemCode: recipeCode,
        itemQuantityRaw: qtyRaw,
        reason: 'Inactive status: ' + (row[statusIdx] || '')
      });
      continue;
    }

    if (!unitCode || !recipeCode || itemQty <= 0) {
      var reasons = [];
      if (!unitCode) reasons.push('missing Unit Product Code');
      if (!recipeCode) reasons.push('missing Recipe Item Code');
      if (itemQty <= 0) reasons.push('invalid Item Quantity (' + String(qtyRaw) + ')');
      skipped.push({
        sheetRow: i + 1,
        unitProductCode: unitCode,
        recipeItemCode: recipeCode,
        itemQuantityRaw: qtyRaw,
        reason: reasons.join('; ')
      });
      continue;
    }

    var line = {
      bomId: bomIdIdx >= 0 ? (row[bomIdIdx] || '').toString().trim() : '',
      bomItemId: bomItemIdx >= 0 ? (row[bomItemIdx] || '').toString().trim() : '',
      recipeItemCode: normalizeProductCodeKey_(recipeCode),
      itemQuantity: itemQty,
      itemQuantityDisplay: (qtyRaw !== null && qtyRaw !== undefined && String(qtyRaw).trim() !== '')
        ? String(qtyRaw).trim()
        : String(itemQty),
      unitProductCode: normalizeProductCodeKey_(unitCode),
      status: statusIdx >= 0 ? (row[statusIdx] || '').toString().trim() : 'Active'
    };
    lines.push(line);
    if (!byUnitProduct[line.unitProductCode]) byUnitProduct[line.unitProductCode] = [];
    byUnitProduct[line.unitProductCode].push(line);
  }

  var byUnitLookup = buildBomUnitLookup_(byUnitProduct);
  return {
    byUnitProduct: byUnitProduct,
    byUnitLookup: byUnitLookup,
    lines: lines,
    totalLines: lines.length,
    skipped: skipped
  };
}

/**
 * Legacy sell → parent buy mapping from Products sheet.
 */
function loadSellToParentMap_() {
  var sellToParent = {};
  var productsSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
  if (!productsSheet) return sellToParent;

  var prodData = productsSheet.getDataRange().getValues();
  if (prodData.length < 2) return sellToParent;

  var headers = prodData[0];
  var typeIdx = headers.findIndex(function(h) { return h && /type/i.test(h.toString()); });
  var codeIdx = headers.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var parentIdx = headers.findIndex(function(h) { return h && /parent.*product.*code|parent.*code/i.test(h.toString()); });
  var unitsIdx = headers.findIndex(function(h) { return h && /units.*per.*parent/i.test(h.toString()); });
  var activeIdx = headers.findIndex(function(h) { return h && /^active$/i.test(h.toString()); });

  for (var i = 1; i < prodData.length; i++) {
    var row = prodData[i];
    if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(row[activeIdx])) continue;
    var type = typeIdx >= 0 ? (row[typeIdx] || '').toString().toLowerCase() : '';
    if (type !== 'sell') continue;
    var code = codeIdx >= 0 ? (row[codeIdx] || '').toString().trim() : '';
    var parentCode = parentIdx >= 0 ? (row[parentIdx] || '').toString().trim() : '';
    var unitsPerParent = unitsIdx >= 0 ? (Number(row[unitsIdx]) || 0) : 0;
    if (!code || !parentCode || unitsPerParent <= 0) continue;
    sellToParent[normalizeProductCodeKey_(code)] = { parentCode: normalizeProductCodeKey_(parentCode), unitsPerParent: unitsPerParent };
  }
  return sellToParent;
}

/**
 * Compute ingredient consumption from sold quantities — BOM-mapped recipes only.
 * Sell products without an active BOM row are skipped (no ParentProduct / DirectBuy fallback).
 */
function computeConsumptionFromSoldMap_(soldMap, ctx) {
  var consumptionMap = {};
  var bomByUnit = (ctx && ctx.bomByUnit) || {};
  var bomLookup = (ctx && ctx.bomLookup) || buildBomUnitLookup_(bomByUnit);

  Object.keys(soldMap || {}).forEach(function(soldCode) {
    var qty = Number(soldMap[soldCode]) || 0;
    if (qty <= 0) return;
    var bomLines = resolveBomLinesForSold_(soldCode, bomLookup);
    if (!bomLines || !bomLines.length) return;

    bomLines.forEach(function(line) {
      var recipeCode = normalizeProductCodeKey_(line.recipeItemCode);
      var perUnit = Number(line.itemQuantity) || 0;
      if (!recipeCode || perUnit <= 0) return;
      consumptionMap[recipeCode] = (consumptionMap[recipeCode] || 0) + (qty * perUnit);
    });
  });

  return consumptionMap;
}

// ========== BOM USAGE MATRIX (base × flavor → grams) ==========
// One row per (base sell product, flavor sell product). Add a row for every flavor
// you offer on Lil Swirl, Big Swirl, Sundaes, etc.

var BOM_USAGE_HEADERS = [
  'Base Unit Product Code',
  'Flavor Unit Product Code',
  'Recipe Item Code',
  'Item Quantity',
  'Status'
];

function ensureBomUsageSheet_() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.BOM_USAGE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.BOM_USAGE);
    sheet.appendRow(BOM_USAGE_HEADERS);
    Logger.log('Created BOM_Usage sheet');
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(BOM_USAGE_HEADERS);
  }
  return sheet;
}

/**
 * @returns {{ byBaseAndFlavor: Object, baseCodes: Object, flavorCodes: Object, lines: Array, skipped: Array }}
 */
function loadActiveBomUsageMatrix_() {
  var byBaseAndFlavor = {};
  var baseCodes = {};
  var flavorCodes = {};
  var lines = [];
  var skipped = [];
  var data = _readRepoValues_('BOM_USAGE');
  if (data.length < 2) {
    return { byBaseAndFlavor: byBaseAndFlavor, baseCodes: baseCodes, flavorCodes: flavorCodes, lines: lines, skipped: skipped };
  }

  var headers = data[0];
  var baseIdx = headers.findIndex(function(h) { return h && /base\s*unit\s*product\s*code/i.test(h.toString()); });
  var flavorIdx = headers.findIndex(function(h) { return h && /flavor\s*unit\s*product\s*code/i.test(h.toString()); });
  var recipeIdx = headers.findIndex(function(h) { return h && /recipe\s*item\s*code/i.test(h.toString()); });
  var qtyIdx = headers.findIndex(function(h) { return h && /item\s*quantity|grams/i.test(h.toString()); });
  var statusIdx = headers.findIndex(function(h) { return h && /^status$/i.test(h.toString()); });

  if (baseIdx === -1 || flavorIdx === -1 || recipeIdx === -1 || qtyIdx === -1) {
    Logger.log('BOM_Usage sheet missing required columns');
    return { byBaseAndFlavor: byBaseAndFlavor, baseCodes: baseCodes, flavorCodes: flavorCodes, lines: lines, skipped: skipped };
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var baseCode = normalizeProductCodeKey_(row[baseIdx]);
    var flavorCode = normalizeProductCodeKey_(row[flavorIdx]);
    var recipeCode = normalizeProductCodeKey_(row[recipeIdx]);
    var qtyRaw = row[qtyIdx];
    var grams = parseBomItemQuantity_(qtyRaw);
    if (!baseCode && !flavorCode && !recipeCode) continue;

    if (statusIdx >= 0 && !isBomRowActive_(row[statusIdx])) {
      skipped.push({ sheetRow: i + 1, baseCode: baseCode, flavorCode: flavorCode, reason: 'Inactive status' });
      continue;
    }
    if (!baseCode || !flavorCode || !recipeCode || grams <= 0) {
      skipped.push({
        sheetRow: i + 1,
        baseCode: baseCode,
        flavorCode: flavorCode,
        reason: 'Missing base/flavor/recipe or invalid grams (' + String(qtyRaw) + ')'
      });
      continue;
    }

    var line = {
      baseUnitProductCode: baseCode,
      flavorUnitProductCode: flavorCode,
      recipeItemCode: recipeCode,
      grams: grams,
      gramsDisplay: String(qtyRaw).trim() || String(grams)
    };
    lines.push(line);
    baseCodes[baseCode] = true;
    flavorCodes[flavorCode] = true;
    byBaseAndFlavor[baseCode + '|' + flavorCode] = line;
  }

  return { byBaseAndFlavor: byBaseAndFlavor, baseCodes: baseCodes, flavorCodes: flavorCodes, lines: lines, skipped: skipped };
}

function loadUnitsForSellingByCode_() {
  var map = {};
  var data = _readRepoValues_('PRODUCTS');
  if (data.length < 2) return map;
  var headers = data[0];
  var codeIdx = headers.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var ufsIdx = headers.findIndex(function(h) { return h && /units\s*for\s*selling|unitsforselling/i.test(h.toString()); });
  if (codeIdx === -1) return map;
  for (var i = 1; i < data.length; i++) {
    var code = normalizeProductCodeKey_(data[i][codeIdx]);
    if (!code) continue;
    var ufs = ufsIdx >= 0 ? (Number(data[i][ufsIdx]) || 0) : 0;
    map[code] = ufs > 0 ? ufs : 1;
  }
  return map;
}

function orderGroupKey_(orderNumber) {
  var s = String(orderNumber || '').trim();
  if (!s) return s;
  var lastDash = s.lastIndexOf('-');
  if (lastDash > 0) {
    var suffix = s.substring(lastDash + 1);
    if (/^\d{1,4}$/.test(suffix)) return s.substring(0, lastDash);
  }
  return s;
}

/**
 * Order lines grouped by checkout (Order Number prefix).
 */
function loadOrderGroupsByDateAndOutlet_(businessDateStr, outlet, options) {
  return readOrderRowsForDate_(businessDateStr, outlet, options).groups;
}

function mergeConsumptionMaps_(target, source) {
  Object.keys(source || {}).forEach(function(code) {
    var qty = Number(source[code]) || 0;
    if (qty <= 0) return;
    target[code] = (target[code] || 0) + qty;
  });
  return target;
}

/**
 * Grams from BOM_Usage: pairs flavor lines with base lines in the same order.
 */
function computeMatrixConsumptionFromOrderGroups_(groups, usageMatrix, meta) {
  var consumptionMap = {};
  var details = [];
  if (!usageMatrix || !usageMatrix.byBaseAndFlavor || !groups || !groups.length) {
    return { consumptionMap: consumptionMap, details: details };
  }

  var ufs = loadUnitsForSellingByCode_();
  var genStr = meta.generatedAtStr || formatTimestampForSheet(new Date());
  var productNames = meta.productNames || {};

  groups.forEach(function(group) {
    var bases = [];
    var flavors = [];
    group.lines.forEach(function(line) {
      var k = normalizeProductCodeKey_(line.code);
      if (usageMatrix.baseCodes[k]) bases.push({ code: k, qty: Number(line.qty) || 0 });
      if (usageMatrix.flavorCodes[k]) flavors.push({ code: k, qty: Number(line.qty) || 0 });
    });
    if (!bases.length || !flavors.length) return;

    var totalWeight = 0;
    bases.forEach(function(b) {
      b.weight = (Number(b.qty) || 0) * (Number(ufs[b.code]) || 1);
      totalWeight += b.weight;
    });
    if (totalWeight <= 0) {
      bases.forEach(function(b) { b.weight = Number(b.qty) || 1; totalWeight += b.weight; });
    }

    flavors.forEach(function(f) {
      var fQty = Number(f.qty) || 0;
      if (fQty <= 0) return;
      bases.forEach(function(b) {
        var share = (Number(b.weight) || 0) / totalWeight;
        var cell = usageMatrix.byBaseAndFlavor[b.code + '|' + f.code];
        if (!cell) return;
        var grams = cell.grams * fQty * share;
        if (grams <= 0) return;
        consumptionMap[cell.recipeItemCode] = (consumptionMap[cell.recipeItemCode] || 0) + grams;
        details.push({
          generatedAt: genStr,
          businessDate: meta.businessDate,
          outlet: meta.outlet,
          sellProductCode: f.code,
          sellProductName: productNames[f.code] || f.code,
          baseProductCode: b.code,
          soldQty: fQty,
          recipeItemCode: cell.recipeItemCode,
          recipeProductName: productNames[cell.recipeItemCode] || cell.recipeItemCode,
          qtyPerUnit: cell.gramsDisplay + 'g × base share ' + roundConsumptionQty_(share),
          consumptionQty: roundConsumptionQty_(grams),
          source: 'BOM_Usage',
          bomId: group.key || ''
        });
      });
    });
  });

  return { consumptionMap: consumptionMap, details: details };
}

/**
 * Standard BOM (packaging, chicken, etc.) + BOM_Usage matrix (base×flavor grams).
 * Flavor SKUs in BOM_Usage are excluded from flat BOM to avoid double counting.
 */
function computeCombinedConsumptionForOutlet_(businessDate, outlet, bomRecipes, usageMatrix, productNames, generatedAt, options) {
  options = options || {};
  var productIndex = options.productIndex || (typeof getCachedProductCodeIndex_ === 'function'
    ? getCachedProductCodeIndex_()
    : null);
  var orderScan = options.orderScan || readOrderRowsForDate_(businessDate, outlet, {
    productIndex: productIndex,
    preferBottomUp: options.preferBottomUp,
    batchSize: options.batchSize,
    strictOutlet: options.strictOutlet
  });
  var soldMap = orderScan.soldMap || {};
  var bomLookup = bomRecipes.byUnitLookup || buildBomUnitLookup_(bomRecipes.byUnitProduct || {}, productIndex);
  var flavorInMatrix = (usageMatrix && usageMatrix.flavorCodes) || {};

  var filteredSold = {};
  Object.keys(soldMap).forEach(function(code) {
    var k = normalizeProductCodeKey_(code);
    if (flavorInMatrix[k]) return;
    filteredSold[k] = soldMap[code];
  });

  var consumptionMap = computeConsumptionFromSoldMap_(filteredSold, {
    bomByUnit: bomRecipes.byUnitProduct,
    bomLookup: bomLookup
  });

  var genStr = formatTimestampForSheet(generatedAt || new Date());
  var standardDetails = buildConsumptionReportDetailLines_(filteredSold, {
    bomByUnit: bomRecipes.byUnitProduct,
    bomLookup: bomLookup
  }, outlet, businessDate, generatedAt, productNames);

  var groups = orderScan.groups || [];
  var matrixResult = computeMatrixConsumptionFromOrderGroups_(groups, usageMatrix, {
    businessDate: businessDate,
    outlet: outlet,
    generatedAtStr: genStr,
    productNames: productNames
  });
  mergeConsumptionMaps_(consumptionMap, matrixResult.consumptionMap);

  var details = standardDetails.concat(matrixResult.details);
  details.sort(function(a, b) {
    return (a.outlet || '').localeCompare(b.outlet || '') ||
      (a.baseProductCode || a.sellProductCode || '').localeCompare(b.baseProductCode || b.sellProductCode || '') ||
      (a.sellProductCode || '').localeCompare(b.sellProductCode || '');
  });

  return { consumptionMap: consumptionMap, details: details, soldMap: soldMap };
}

function writeBomUsageAuditToSheet_(bomAuditSheet, usageMatrix, bomRecipes) {
  if (!bomAuditSheet || !usageMatrix) return;
  var rows = [];
  (usageMatrix.lines || []).forEach(function(line) {
    rows.push([
      'Usage Loaded',
      '',
      line.baseUnitProductCode,
      line.flavorUnitProductCode + ' → ' + line.recipeItemCode,
      line.grams,
      line.gramsDisplay,
      'Active',
      'OK',
      'Grams per 1 flavor unit with this base in same order'
    ]);
  });
  (usageMatrix.skipped || []).forEach(function(s) {
    rows.push(['Usage Skipped', s.sheetRow || '', s.baseCode || '', s.flavorCode || '', '', '', '', 'SKIPPED', s.reason || '']);
  });
  Object.keys(usageMatrix.baseCodes || {}).forEach(function(base) {
    var flavorList = (usageMatrix.lines || []).filter(function(l) { return l.baseUnitProductCode === base; })
      .map(function(l) { return l.flavorUnitProductCode; }).join(', ');
    rows.push(['Usage Base Summary', '', base, flavorList, '', '', '', 'INFO',
      'Add one row per flavor (any flavor SKU) used with this base']);
  });
  if (!rows.length) return;
  if (typeof appendRowsAtomic_ === 'function') {
    appendRowsAtomic_(bomAuditSheet, rows);
  } else {
    rows.forEach(function(r) { bomAuditSheet.appendRow(r); });
  }
}

/**
 * Single Orders scan for sold qty + order groups (optional bottom-up for append-only sheets).
 * @returns {{ soldMap: Object, groups: Array, rowsScanned: number, scanMode: string }}
 */
function readOrderRowsForDate_(businessDateStr, outlet, options) {
  options = options || {};
  var soldMap = {};
  var groups = {};
  var batchSize = options.batchSize || (typeof CONSUMPTION_ORDER_BATCH_SIZE !== 'undefined'
    ? CONSUMPTION_ORDER_BATCH_SIZE
    : 500);
  var preferBottomUp = options.preferBottomUp !== false;
  var firstBatch = preferBottomUp ? Repo.readRowsFromEnd('ORDERS', 0, batchSize) : null;
  var fullOrderData = preferBottomUp ? null : _readRepoValues_('ORDERS');
  var headerRow = preferBottomUp ? firstBatch.headers : (fullOrderData[0] || []);
  var totalRows = preferBottomUp ? firstBatch.totalRows : Math.max(0, fullOrderData.length - 1);
  if (!headerRow.length || !totalRows) {
    return { soldMap: soldMap, groups: [], rowsScanned: 0, scanMode: 'empty' };
  }

  var dateIdx = headerRow.findIndex(function(h) { return h && /timestamp|date/i.test(h.toString()); });
  var orderNumIdx = headerRow.findIndex(function(h) { return h && /order\s*number/i.test(h.toString()); });
  var codeIdx = headerRow.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var nameIdx = headerRow.findIndex(function(h) { return h && /product.*name|^name$/i.test(h.toString()); });
  var qtyIdx = headerRow.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var outletIdx = headerRow.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  if (options.strictOutlet && (outletIdx < 0 || dateIdx < 0)) {
    throw new Error('Orders is missing a date or outlet column.');
  }
  var productIndex = options.productIndex || (typeof getCachedProductCodeIndex_ === 'function'
    ? getCachedProductCodeIndex_()
    : loadProductCodeIndex_());

  var rowsScanned = 0;
  var scanMode = preferBottomUp ? 'bottom_up' : 'full';

  function processMatchingRow(row, rowNum) {
    var soldCode = resolveSoldCodeFromOrderRow_(row, codeIdx, nameIdx, productIndex);
    var qty = qtyIdx >= 0 ? (Number(row[qtyIdx]) || 0) : 0;
    if (!soldCode || qty <= 0) return;

    soldMap[soldCode] = (soldMap[soldCode] || 0) + qty;
    var orderNum = orderNumIdx >= 0 ? String(row[orderNumIdx] || '').trim() : '';
    var gKey = orderGroupKey_(orderNum) || ('ROW-' + rowNum);
    if (!groups[gKey]) groups[gKey] = { key: gKey, orderNumber: orderNum, lines: [] };
    groups[gKey].lines.push({ code: soldCode, qty: qty });
  }

  function rowMatchesOutlet_(row) {
    if (!outlet || outletIdx === -1) return true;
    if (options.strictOutlet) return canonicalOutletKey_(row[outletIdx]) === canonicalOutletKey_(outlet);
    if (typeof outletRowMatchesUser_ === 'function') {
      return outletRowMatchesUser_(row[outletIdx], outlet);
    }
    return String(row[outletIdx] || '').trim() === String(outlet).trim();
  }

  function rowMatchesDate_(row) {
    if (dateIdx < 0 || !row[dateIdx]) return false;
    var orderDate = row[dateIdx] instanceof Date ? row[dateIdx] : new Date(row[dateIdx]);
    if (isNaN(orderDate.getTime())) return false;
    if (typeof getAppDateKey_ === 'function') {
      return getAppDateKey_(orderDate) === businessDateStr;
    }
    return formatDateForComparison(orderDate) === businessDateStr;
  }

  function rowDateKey_(row) {
    if (dateIdx < 0 || !row[dateIdx]) return null;
    var orderDate = row[dateIdx] instanceof Date ? row[dateIdx] : new Date(row[dateIdx]);
    if (isNaN(orderDate.getTime())) return null;
    if (typeof getAppDateKey_ === 'function') {
      return getAppDateKey_(orderDate);
    }
    return formatDateForComparison(orderDate);
  }

  if (preferBottomUp) {
    var offset = 0;
    while (offset < firstBatch.totalRows) {
      var batchInfo = offset === 0
        ? firstBatch
        : Repo.readRowsFromEnd('ORDERS', offset, batchSize, firstBatch.totalRows);
      var batch = batchInfo.rows.map(function (row) {
        return headerRow.map(function (header) { return row[header]; });
      });
      for (var i = batch.length - 1; i >= 0; i--) {
        rowsScanned++;
        var row = batch[i];
        var rowNum = batchInfo.startRow + i;
        var dateKey = rowDateKey_(row);
        if (!dateKey) continue;
        if (dateKey < businessDateStr) {
          return {
            soldMap: soldMap,
            groups: Object.keys(groups).map(function(k) { return groups[k]; }),
            rowsScanned: rowsScanned,
            scanMode: scanMode
          };
        }
        if (dateKey !== businessDateStr) continue;
        if (!rowMatchesOutlet_(row)) continue;
        processMatchingRow(row, rowNum);
      }
      offset += batch.length;
    }
  } else {
    scanMode = 'full';
    for (var fr = 1; fr < fullOrderData.length; fr++) {
      rowsScanned++;
      var fullRow = fullOrderData[fr];
      if (!rowMatchesDate_(fullRow)) continue;
      if (!rowMatchesOutlet_(fullRow)) continue;
      processMatchingRow(fullRow, fr + 1);
    }
  }

  return {
    soldMap: soldMap,
    groups: Object.keys(groups).map(function(k) { return groups[k]; }),
    rowsScanned: rowsScanned,
    scanMode: scanMode
  };
}

function getSoldQtyByDateAndOutlet_(businessDateStr, outlet, options) {
  return readOrderRowsForDate_(businessDateStr, outlet, options).soldMap;
}

/**
 * All order rows for one outlet (no business-date filter) — for lifetime BOM stock.
 * @returns {{ soldMap: Object, groups: Array, rowsScanned: number, scanMode: string }}
 */
function readAllOrderRowsForOutlet_(outlet, options) {
  options = options || {};
  var soldMap = {};
  var groups = {};
  var orderData = _readRepoValues_('ORDERS');
  if (orderData.length < 2) {
    return { soldMap: soldMap, groups: [], rowsScanned: 0, scanMode: 'empty' };
  }

  var headerRow = orderData[0];
  var orderNumIdx = headerRow.findIndex(function(h) { return h && /order\s*number/i.test(h.toString()); });
  var codeIdx = headerRow.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var nameIdx = headerRow.findIndex(function(h) { return h && /product.*name|^name$/i.test(h.toString()); });
  var qtyIdx = headerRow.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var outletIdx = headerRow.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  var productIndex = options.productIndex || (typeof getCachedProductCodeIndex_ === 'function'
    ? getCachedProductCodeIndex_()
    : loadProductCodeIndex_());
  var rowsScanned = 0;

  function processMatchingRow(row, rowNum) {
    var soldCode = resolveSoldCodeFromOrderRow_(row, codeIdx, nameIdx, productIndex);
    var qty = qtyIdx >= 0 ? (Number(row[qtyIdx]) || 0) : 0;
    if (!soldCode || qty <= 0) return;

    soldMap[soldCode] = (soldMap[soldCode] || 0) + qty;
    var orderNum = orderNumIdx >= 0 ? String(row[orderNumIdx] || '').trim() : '';
    var gKey = orderGroupKey_(orderNum) || ('ROW-' + rowNum);
    if (!groups[gKey]) groups[gKey] = { key: gKey, orderNumber: orderNum, lines: [] };
    groups[gKey].lines.push({ code: soldCode, qty: qty });
  }

  function rowMatchesOutlet_(row) {
    if (!outlet || outletIdx === -1) return true;
    if (typeof outletRowMatchesUser_ === 'function') {
      return outletRowMatchesUser_(row[outletIdx], outlet);
    }
    return String(row[outletIdx] || '').trim() === String(outlet).trim();
  }

  for (var fr = 1; fr < orderData.length; fr++) {
    rowsScanned++;
    var fullRow = orderData[fr];
    if (!rowMatchesOutlet_(fullRow)) continue;
    processMatchingRow(fullRow, fr + 1);
  }

  return {
    soldMap: soldMap,
    groups: Object.keys(groups).map(function(k) { return groups[k]; }),
    rowsScanned: rowsScanned,
    scanMode: 'lifetime'
  };
}

/**
 * One Orders read — pivot sold qty by product code for rows on or before throughDateStr.
 * @returns {{ soldMap: Object, groups: Array, rowsScanned: number, scanMode: string }}
 */
function readOrderSoldMapThroughDateForOutlet_(outlet, throughDateStr, options) {
  options = options || {};
  var soldMap = {};
  var groups = {};
  var orderData = _readRepoValues_('ORDERS');
  if (orderData.length < 2) {
    return { soldMap: soldMap, groups: [], rowsScanned: 0, scanMode: 'empty' };
  }

  var headerRow = orderData[0];
  var dateIdx = headerRow.findIndex(function(h) { return h && /timestamp|date/i.test(h.toString()); });
  var orderNumIdx = headerRow.findIndex(function(h) { return h && /order\s*number/i.test(h.toString()); });
  var codeIdx = headerRow.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
  var nameIdx = headerRow.findIndex(function(h) { return h && /product.*name|^name$/i.test(h.toString()); });
  var qtyIdx = headerRow.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
  var outletIdx = headerRow.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
  var productIndex = options.productIndex || (typeof getCachedProductCodeIndex_ === 'function'
    ? getCachedProductCodeIndex_()
    : loadProductCodeIndex_());
  var rowsScanned = 0;
  var throughKey = String(throughDateStr || '').trim();

  function rowMatchesOutlet_(row) {
    if (!outlet || outletIdx === -1) return true;
    if (typeof outletRowMatchesUser_ === 'function') {
      return outletRowMatchesUser_(row[outletIdx], outlet);
    }
    return String(row[outletIdx] || '').trim() === String(outlet).trim();
  }

  function rowDateKey_(row) {
    if (dateIdx < 0 || !row[dateIdx]) return null;
    var orderDate = row[dateIdx] instanceof Date ? row[dateIdx] : new Date(row[dateIdx]);
    if (isNaN(orderDate.getTime())) return null;
    return typeof getAppDateKey_ === 'function'
      ? getAppDateKey_(orderDate)
      : formatDateForComparison(orderDate);
  }

  function processMatchingRow(row, rowNum) {
    var soldCode = resolveSoldCodeFromOrderRow_(row, codeIdx, nameIdx, productIndex);
    var qty = qtyIdx >= 0 ? (Number(row[qtyIdx]) || 0) : 0;
    if (!soldCode || qty <= 0) return;

    soldMap[soldCode] = (soldMap[soldCode] || 0) + qty;
    var orderNum = orderNumIdx >= 0 ? String(row[orderNumIdx] || '').trim() : '';
    var gKey = orderGroupKey_(orderNum) || ('ROW-' + rowNum);
    if (!groups[gKey]) groups[gKey] = { key: gKey, orderNumber: orderNum, lines: [] };
    groups[gKey].lines.push({ code: soldCode, qty: qty });
  }

  for (var fr = 1; fr < orderData.length; fr++) {
    rowsScanned++;
    var fullRow = orderData[fr];
    if (!rowMatchesOutlet_(fullRow)) continue;
    var dateKey = rowDateKey_(fullRow);
    if (!dateKey || !throughKey || dateKey > throughKey) continue;
    processMatchingRow(fullRow, fr + 1);
  }

  return {
    soldMap: soldMap,
    groups: Object.keys(groups).map(function(k) { return groups[k]; }),
    rowsScanned: rowsScanned,
    scanMode: 'through_date'
  };
}

/**
 * BOM ingredient consumption from orders on or before throughDateStr (one Orders scan).
 */
function computeBomConsumptionThroughDateForOutlet_(outlet, throughDateStr, options) {
  options = options || {};
  var bomRecipes = options.bomRecipes || loadActiveBomRecipes_();
  var usageMatrix = options.usageMatrix || loadActiveBomUsageMatrix_();
  var productNames = options.productNames || (typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {});
  var orderScan = options.orderScan || readOrderSoldMapThroughDateForOutlet_(outlet, throughDateStr, {
    productIndex: options.productIndex
  });
  var soldMap = orderScan.soldMap || {};
  var bomLookup = bomRecipes.byUnitLookup || buildBomUnitLookup_(bomRecipes.byUnitProduct || {}, options.productIndex);
  var flavorInMatrix = (usageMatrix && usageMatrix.flavorCodes) || {};

  var filteredSold = {};
  Object.keys(soldMap).forEach(function(code) {
    var k = normalizeProductCodeKey_(code);
    if (flavorInMatrix[k]) return;
    filteredSold[k] = soldMap[code];
  });

  var consumptionMap = computeConsumptionFromSoldMap_(filteredSold, {
    bomByUnit: bomRecipes.byUnitProduct,
    bomLookup: bomLookup
  });

  var matrixResult = computeMatrixConsumptionFromOrderGroups_(orderScan.groups || [], usageMatrix, {
    businessDate: throughDateStr,
    outlet: outlet,
    generatedAtStr: formatTimestampForSheet(new Date()),
    productNames: productNames
  });
  mergeConsumptionMaps_(consumptionMap, matrixResult.consumptionMap);

  return {
    consumptionMap: consumptionMap,
    rowsScanned: orderScan.rowsScanned || 0
  };
}

/**
 * Active recipe item codes from BOM + BOM_Usage sheets.
 */
function buildRecipeItemCodeSet_(bomRecipes, usageMatrix) {
  var codes = {};
  (bomRecipes && bomRecipes.lines || []).forEach(function(line) {
    var code = normalizeProductCodeKey_(line.recipeItemCode);
    if (code) codes[code] = true;
  });
  (usageMatrix && usageMatrix.lines || []).forEach(function(line) {
    var code = normalizeProductCodeKey_(line.recipeItemCode);
    if (code) codes[code] = true;
  });
  return codes;
}

/**
 * Lifetime ingredient consumption for one outlet (standard BOM + BOM_Usage matrix).
 */
function computeLifetimeBomConsumptionForOutlet_(outlet, options) {
  options = options || {};
  var bomRecipes = options.bomRecipes || loadActiveBomRecipes_();
  var usageMatrix = options.usageMatrix || loadActiveBomUsageMatrix_();
  var productNames = options.productNames || (typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {});
  var orderScan = options.orderScan || readAllOrderRowsForOutlet_(outlet, {
    productIndex: options.productIndex
  });
  var soldMap = orderScan.soldMap || {};
  var bomLookup = bomRecipes.byUnitLookup || buildBomUnitLookup_(bomRecipes.byUnitProduct || {}, options.productIndex);
  var flavorInMatrix = (usageMatrix && usageMatrix.flavorCodes) || {};

  var filteredSold = {};
  Object.keys(soldMap).forEach(function(code) {
    var k = normalizeProductCodeKey_(code);
    if (flavorInMatrix[k]) return;
    filteredSold[k] = soldMap[code];
  });

  var consumptionMap = computeConsumptionFromSoldMap_(filteredSold, {
    bomByUnit: bomRecipes.byUnitProduct,
    bomLookup: bomLookup
  });

  var matrixResult = computeMatrixConsumptionFromOrderGroups_(orderScan.groups || [], usageMatrix, {
    businessDate: '',
    outlet: outlet,
    generatedAtStr: formatTimestampForSheet(new Date()),
    productNames: productNames
  });
  mergeConsumptionMaps_(consumptionMap, matrixResult.consumptionMap);

  return {
    consumptionMap: consumptionMap,
    recipeItemCodes: buildRecipeItemCodeSet_(bomRecipes, usageMatrix),
    rowsScanned: orderScan.rowsScanned || 0
  };
}

/**
 * Outlets to process: configured Outlets sheet + distinct order outlets.
 */
function getDistinctOutletsForConsumption_() {
  var outletSet = {};
  Repo.readRows('OUTLETS').forEach(function (row) {
    var keys = Object.keys(row);
    var outletName = '';
    var outletId = '';
    keys.forEach(function (key) {
      var normalizedKey = key.trim();
      if (/\bname\b/i.test(normalizedKey)) outletName = outletName || row[key];
      else if (/\bid\b/i.test(normalizedKey)) outletId = outletId || row[key];
      else if (/outlet/i.test(normalizedKey)) outletName = outletName || row[key];
    });
    var name = String(outletName || outletId || '').trim();
    if (name) outletSet[name] = true;
  });

  ['ORDERS', 'PURCHASES', 'INVENTORY_CONSUMPTION'].forEach(function (collection) {
    Repo.readRows(collection).forEach(function (row) {
      var keys = Object.keys(row);
      for (var i = 0; i < keys.length; i++) {
        if (!/^outlet$/i.test(keys[i].trim())) continue;
        var outlet = String(row[keys[i]] || '').trim();
        if (outlet) outletSet[outlet] = true;
        break;
      }
    });
  });

  return Object.keys(outletSet).sort();
}

function getBomAdminSummary_() {
  var recipes = loadActiveBomRecipes_();
  return {
    sheetName: SHEETS.BOM,
    headers: BOM_HEADERS,
    totalLines: recipes.totalLines,
    unitProductCount: Object.keys(recipes.byUnitProduct).length,
    lines: recipes.lines.slice(0, 500)
  };
}

function getBomForAdminSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  return { success: true, data: getBomAdminSummary_() };
}

// ========== CONSUMPTION VERIFICATION REPORT (SHEET) ==========

var CONSUMPTION_REPORT_SUMMARY_HEADERS = [
  'Generated At',
  'Business Date',
  'Outlet',
  'Recipe Item Code',
  'Recipe Product Name',
  'Total Consumption Qty',
  'Detail Lines'
];

var CONSUMPTION_REPORT_DETAIL_HEADERS = [
  'Generated At',
  'Business Date',
  'Outlet',
  'Sell Product Code',
  'Sell Product Name',
  'Base Product Code',
  'Sold Qty',
  'Recipe Item Code',
  'Recipe Product Name',
  'Qty Per Unit Sold',
  'Consumption Qty',
  'Source',
  'BOM Id'
];

var CONSUMPTION_REPORT_BOM_HEADERS = [
  'Row Type',
  'BOM Sheet Row',
  'Unit Product Code',
  'Recipe Item Code',
  'Item Quantity (parsed)',
  'Item Quantity (raw)',
  'Status',
  'Load Result',
  'Notes'
];

function ensureConsumptionReportSheets_() {
  var ss = getGlobalSpreadsheet();
  var summary = ss.getSheetByName(SHEETS.CONSUMPTION_REPORT);
  if (!summary) {
    summary = ss.insertSheet(SHEETS.CONSUMPTION_REPORT);
    summary.appendRow(CONSUMPTION_REPORT_SUMMARY_HEADERS);
    Logger.log('Created Consumption_Report sheet');
  } else if (summary.getLastRow() === 0) {
    summary.appendRow(CONSUMPTION_REPORT_SUMMARY_HEADERS);
  }

  var detail = ss.getSheetByName(SHEETS.CONSUMPTION_REPORT_DETAIL);
  if (!detail) {
    detail = ss.insertSheet(SHEETS.CONSUMPTION_REPORT_DETAIL);
    detail.appendRow(CONSUMPTION_REPORT_DETAIL_HEADERS);
    Logger.log('Created Consumption_Report_Detail sheet');
  } else if (detail.getLastRow() === 0) {
    detail.appendRow(CONSUMPTION_REPORT_DETAIL_HEADERS);
  }

  var bomAudit = ss.getSheetByName(SHEETS.CONSUMPTION_REPORT_BOM);
  if (!bomAudit) {
    bomAudit = ss.insertSheet(SHEETS.CONSUMPTION_REPORT_BOM);
    bomAudit.appendRow(CONSUMPTION_REPORT_BOM_HEADERS);
    Logger.log('Created Consumption_Report_BOM sheet');
  } else if (bomAudit.getLastRow() === 0) {
    bomAudit.appendRow(CONSUMPTION_REPORT_BOM_HEADERS);
  }

  return { summary: summary, detail: detail, bomAudit: bomAudit };
}

function writeBomAuditToSheet_(bomAuditSheet, bomRecipes, generatedAt) {
  if (!bomAuditSheet) return;
  var genStr = formatTimestampForSheet(generatedAt || new Date());
  var rows = [];

  (bomRecipes.lines || []).forEach(function(line) {
    rows.push([
      'Loaded',
      '',
      line.unitProductCode,
      line.recipeItemCode,
      line.itemQuantity,
      line.itemQuantityDisplay,
      line.status || 'Active',
      'OK',
      'Used in consumption calc'
    ]);
  });

  (bomRecipes.skipped || []).forEach(function(s) {
    rows.push([
      'Skipped',
      s.sheetRow || '',
      s.unitProductCode || '',
      s.recipeItemCode || '',
      '',
      s.itemQuantityRaw != null ? String(s.itemQuantityRaw) : '',
      '',
      'SKIPPED',
      s.reason || ''
    ]);
  });

  Object.keys(bomRecipes.byUnitProduct || {}).forEach(function(unitCode) {
    var unitLines = bomRecipes.byUnitProduct[unitCode] || [];
    var recipeList = unitLines.map(function(l) { return l.recipeItemCode; }).join(', ');
    rows.push([
      'Unit Summary',
      '',
      unitCode,
      recipeList,
      unitLines.length,
      '',
      '',
      'INFO',
      unitLines.length + ' ingredient line(s) for this sell product'
    ]);
  });

  if (bomAuditSheet.getLastRow() > 1) {
    bomAuditSheet.deleteRows(2, bomAuditSheet.getLastRow() - 1);
  }
  if (!rows.length) return;
  if (typeof appendRowsAtomic_ === 'function') {
    appendRowsAtomic_(bomAuditSheet, rows);
  } else {
    rows.forEach(function(r) { bomAuditSheet.appendRow(r); });
  }
}

function roundConsumptionQty_(n) {
  return Math.round((Number(n) || 0) * 1000000) / 1000000;
}

/**
 * Line-by-line consumption audit — BOM-mapped recipes only.
 */
function buildConsumptionReportDetailLines_(soldMap, ctx, outlet, businessDate, generatedAt, productNames) {
  var details = [];
  var bomLookup = (ctx && ctx.bomLookup) || buildBomUnitLookup_((ctx && ctx.bomByUnit) || {});
  var genStr = formatTimestampForSheet(generatedAt || new Date());
  productNames = productNames || {};

  Object.keys(soldMap || {}).forEach(function(sellCode) {
    var soldQty = Number(soldMap[sellCode]) || 0;
    if (soldQty <= 0) return;
    var sellKey = normalizeProductCodeKey_(sellCode);
    var bomLines = resolveBomLinesForSold_(sellCode, bomLookup);
    if (!bomLines || !bomLines.length) return;

    bomLines.forEach(function(line) {
      var perUnit = Number(line.itemQuantity) || 0;
      if (!line.recipeItemCode || perUnit <= 0) return;
      details.push({
        generatedAt: genStr,
        businessDate: businessDate,
        outlet: outlet,
        sellProductCode: sellKey,
        sellProductName: productNames[sellKey] || productNames[sellCode] || sellCode,
        baseProductCode: '',
        soldQty: soldQty,
        recipeItemCode: line.recipeItemCode,
        recipeProductName: productNames[line.recipeItemCode] || line.recipeItemCode,
        qtyPerUnit: line.itemQuantityDisplay || String(perUnit),
        consumptionQty: roundConsumptionQty_(soldQty * perUnit),
        source: 'BOM',
        bomId: line.bomId || ''
      });
    });
  });

  details.sort(function(a, b) {
    return (a.outlet || '').localeCompare(b.outlet || '') ||
      (a.sellProductCode || '').localeCompare(b.sellProductCode || '') ||
      (a.recipeItemCode || '').localeCompare(b.recipeItemCode || '');
  });
  return details;
}

/**
 * Expand detail lines into one prefill row per (recipeItemCode + sellProductCode).
 */
function buildConsumptionPrefillLinesFromDetails_(details, buyByCode, productNames) {
  buyByCode = buyByCode || {};
  productNames = productNames || {};
  var groups = {};

  (details || []).forEach(function(d) {
    var recipe = normalizeProductCodeKey_(d.recipeItemCode);
    if (!recipe) return;
    var sellKey = normalizeProductCodeKey_(d.sellProductCode || d.baseProductCode || '');
    var groupKey = recipe + '|' + sellKey;
    if (!groups[groupKey]) {
      groups[groupKey] = {
        recipeItemCode: recipe,
        sellProductCode: sellKey,
        sellProductName: d.sellProductName || productNames[sellKey] || sellKey,
        consumptionQty: 0,
        soldQty: 0,
        perUnitSeen: {}
      };
    }
    var g = groups[groupKey];
    g.consumptionQty += Number(d.consumptionQty) || 0;
    var soldQty = Number(d.soldQty) || 0;
    if (soldQty > g.soldQty) g.soldQty = soldQty;
    var perDisplay = d.qtyPerUnit != null ? String(d.qtyPerUnit).trim() : '';
    if (perDisplay) g.perUnitSeen[perDisplay] = true;
  });

  var items = [];
  Object.keys(groups).forEach(function(groupKey) {
    var g = groups[groupKey];
    var suggested = roundConsumptionQty_(g.consumptionQty);
    if (suggested <= 0) return;
    var code = g.recipeItemCode;
    var p = buyByCode[code];
    var perUnits = Object.keys(g.perUnitSeen);
    items.push({
      productCode: code,
      sellProductCode: g.sellProductCode,
      sellProductName: g.sellProductName,
      name: p ? p.name : (productNames[code] || code),
      category: p ? p.category : '',
      navCategory: p ? p.navCategory : '',
      unit: p ? p.unit : '',
      buyPrice: p ? p.buyPrice : 0,
      suggestedQuantity: Math.round(suggested * 1000) / 1000,
      soldUnitsToday: roundConsumptionQty_(g.soldQty),
      qtyPerServing: perUnits.length === 1 ? perUnits[0] : (perUnits.length > 1 ? perUnits.join(' / ') : '')
    });
  });

  items.sort(function(a, b) {
    var nameCmp = (a.name || '').localeCompare(b.name || '');
    if (nameCmp !== 0) return nameCmp;
    return (a.sellProductName || a.sellProductCode || '').localeCompare(b.sellProductName || b.sellProductCode || '');
  });
  return items;
}

/**
 * Per buy/recipe item: total sell units today + BOM qty-per-serving display(s).
 */
function aggregateConsumptionDetailMeta_(details) {
  var meta = {};
  (details || []).forEach(function(d) {
    var recipe = normalizeProductCodeKey_(d.recipeItemCode);
    if (!recipe) return;
    if (!meta[recipe]) {
      meta[recipe] = { soldBySell: {}, perUnitSeen: {} };
    }
    var sellKey = normalizeProductCodeKey_(d.sellProductCode || d.baseProductCode || '');
    var soldQty = Number(d.soldQty) || 0;
    if (sellKey && soldQty > 0) {
      meta[recipe].soldBySell[sellKey] = Math.max(meta[recipe].soldBySell[sellKey] || 0, soldQty);
    }
    var perDisplay = d.qtyPerUnit != null ? String(d.qtyPerUnit).trim() : '';
    if (perDisplay) meta[recipe].perUnitSeen[perDisplay] = true;
  });

  var result = {};
  Object.keys(meta).forEach(function(recipe) {
    var m = meta[recipe];
    var soldUnitsToday = 0;
    Object.keys(m.soldBySell).forEach(function(k) {
      soldUnitsToday += m.soldBySell[k];
    });
    var perUnits = Object.keys(m.perUnitSeen);
    result[recipe] = {
      soldUnitsToday: roundConsumptionQty_(soldUnitsToday),
      qtyPerServing: perUnits.length === 1 ? perUnits[0] : (perUnits.length > 1 ? perUnits.join(' / ') : '')
    };
  });
  return result;
}

function buildConsumptionReportSummaryLines_(detailLines) {
  var map = {};
  (detailLines || []).forEach(function(d) {
    var key = (d.outlet || '') + '|' + (d.recipeItemCode || '');
    if (!map[key]) {
      map[key] = {
        generatedAt: d.generatedAt,
        businessDate: d.businessDate,
        outlet: d.outlet,
        recipeItemCode: d.recipeItemCode,
        recipeProductName: d.recipeProductName,
        totalQty: 0,
        detailLines: 0
      };
    }
    map[key].totalQty += Number(d.consumptionQty) || 0;
    map[key].detailLines++;
  });

  return Object.keys(map).map(function(k) {
    var row = map[k];
    return {
      generatedAt: row.generatedAt,
      businessDate: row.businessDate,
      outlet: row.outlet,
      recipeItemCode: row.recipeItemCode,
      recipeProductName: row.recipeProductName,
      totalQty: roundConsumptionQty_(row.totalQty),
      detailLines: row.detailLines
    };
  }).sort(function(a, b) {
    return (a.outlet || '').localeCompare(b.outlet || '') ||
      (a.recipeProductName || '').localeCompare(b.recipeProductName || '');
  });
}

/**
 * Write full consumption verification report to Consumption_Report + Consumption_Report_Detail.
 * Does NOT post to Inventory_Consumption.
 * @param {string} [businessDateStr] - YYYY-MM-DD
 * @param {string} [onlyOutlet] - optional outlet filter
 */
function generateConsumptionVerificationReport(businessDateStr, onlyOutlet) {
  if (typeof ensureBomSheet_ === 'function') ensureBomSheet_();
  if (typeof ensureBomUsageSheet_ === 'function') ensureBomUsageSheet_();
  var sheets = ensureConsumptionReportSheets_();

  var businessDate = businessDateStr || formatDateForComparison(new Date());
  var generatedAt = new Date();
  var genStr = formatTimestampForSheet(generatedAt);

  var bomRecipes = loadActiveBomRecipes_();
  var usageMatrix = loadActiveBomUsageMatrix_();
  var productNames = typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {};

  writeBomAuditToSheet_(sheets.bomAudit, bomRecipes, generatedAt);
  writeBomUsageAuditToSheet_(sheets.bomAudit, usageMatrix, bomRecipes);

  var outlets = getDistinctOutletsForConsumption_();
  if (onlyOutlet) outlets = [String(onlyOutlet).trim()];

  var allDetails = [];
  outlets.forEach(function(outlet) {
    var result = computeCombinedConsumptionForOutlet_(
      businessDate, outlet, bomRecipes, usageMatrix, productNames, generatedAt
    );
    allDetails = allDetails.concat(result.details || []);
  });

  var summaryLines = buildConsumptionReportSummaryLines_(allDetails);

  // Replace prior report rows (keep headers)
  if (sheets.summary.getLastRow() > 1) {
    sheets.summary.deleteRows(2, sheets.summary.getLastRow() - 1);
  }
  if (sheets.detail.getLastRow() > 1) {
    sheets.detail.deleteRows(2, sheets.detail.getLastRow() - 1);
  }

  var summaryRows = summaryLines.map(function(s) {
    return [
      s.generatedAt || genStr,
      s.businessDate || businessDate,
      s.outlet || '',
      s.recipeItemCode || '',
      s.recipeProductName || '',
      s.totalQty || 0,
      s.detailLines || 0
    ];
  });

  var detailRows = allDetails.map(function(d) {
    return [
      d.generatedAt || genStr,
      d.businessDate || businessDate,
      d.outlet || '',
      d.sellProductCode || '',
      d.sellProductName || '',
      d.baseProductCode || '',
      d.soldQty || 0,
      d.recipeItemCode || '',
      d.recipeProductName || '',
      d.qtyPerUnit || '',
      d.consumptionQty || 0,
      d.source || '',
      d.bomId || ''
    ];
  });

  if (summaryRows.length && typeof appendRowsAtomic_ === 'function') {
    appendRowsAtomic_(sheets.summary, summaryRows);
  } else if (summaryRows.length) {
    summaryRows.forEach(function(r) { sheets.summary.appendRow(r); });
  }

  if (detailRows.length && typeof appendRowsAtomic_ === 'function') {
    appendRowsAtomic_(sheets.detail, detailRows);
  } else if (detailRows.length) {
    detailRows.forEach(function(r) { sheets.detail.appendRow(r); });
  }

  return {
    success: true,
    businessDate: businessDate,
    generatedAt: genStr,
    summarySheet: SHEETS.CONSUMPTION_REPORT,
    detailSheet: SHEETS.CONSUMPTION_REPORT_DETAIL,
    bomAuditSheet: SHEETS.CONSUMPTION_REPORT_BOM,
    bomSkippedRows: (bomRecipes.skipped || []).length,
    sayapIngredientCount: (bomRecipes.byUnitProduct && bomRecipes.byUnitProduct['AYAM_AYAM_ORI_SAYAP'])
      ? bomRecipes.byUnitProduct['AYAM_AYAM_ORI_SAYAP'].length
      : 0,
    summaryRows: summaryRows.length,
    detailRows: detailRows.length,
    outlets: outlets,
    message: 'BOM report: ' + summaryRows.length + ' summary row(s), ' +
      detailRows.length + ' detail row(s), ' +
      (usageMatrix.lines || []).length + ' usage matrix row(s) for ' + businessDate
  };
}

function generateConsumptionVerificationReportSecure(params) {
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  var businessDate = params && params.businessDate ? String(params.businessDate).trim() : '';
  var outlet = params && params.outlet ? String(params.outlet).trim() : '';
  return generateConsumptionVerificationReport(businessDate || undefined, outlet || undefined);
}
