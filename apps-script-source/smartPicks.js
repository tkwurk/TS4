// ========== SMART PICKS - Cached Top Products ==========
// Pre-computed top 20 products stored in SmartPicks sheet.
// Refreshed by scheduled trigger (midnight) or admin ad-hoc.

/**
 * Ensure SmartPicks sheet exists with proper headers.
 */
function ensureSmartPicksSheet() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SMART_PICKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.SMART_PICKS);
    sheet.appendRow(['Type', 'Outlet', 'Supplier', 'NavCategory', 'ProductCode', 'Rank', 'UpdatedAt']);
    Logger.log('Created SmartPicks sheet');
  } else {
    // Migrate: ensure correct 7-column format
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colCount = headerRow.length;
    if (colCount < 6) {
      // Very old 5-col format: insert Supplier + NavCategory
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('Supplier');
      sheet.insertColumnAfter(3);
      sheet.getRange(1, 4).setValue('NavCategory');
      Logger.log('SmartPicks: migrated from 5-col to 7-col format');
    } else if (colCount === 6) {
      // 6-col format (has Supplier, missing NavCategory): insert NavCategory after Supplier
      var col3 = String(headerRow[2] || '').trim().toLowerCase();
      if (col3 === 'supplier') {
        sheet.insertColumnAfter(3);
        sheet.getRange(1, 4).setValue('NavCategory');
        Logger.log('SmartPicks: migrated from 6-col to 7-col format (added NavCategory)');
      }
    }
    // 7-col format already correct
  }
  return sheet;
}

/**
 * Build a set of valid product codes/names by type from Products sheet.
 * @param {string} type - 'Sell' or 'Buy'
 * @returns {Object} - { code: true } lookup
 */
function _getProductCodesByType(type) {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var typeIdx = headers.findIndex(function(h) { return h === 'type' || h === 'product type'; });
  var codeIdx = headers.findIndex(function(h) { return /product.*code|code/i.test(h); });
  var nameIdx = headers.findIndex(function(h) { return h === 'name' || h === 'product name' || h === 'product'; });

  var activeIdx = headers.findIndex(function(h) { return h === 'active'; });

  Logger.log('SmartPicks _getProductCodesByType(' + type + '): typeIdx=' + typeIdx + ', codeIdx=' + codeIdx + ', nameIdx=' + nameIdx + ', headers=' + JSON.stringify(headers));

  var lookup = {};
  for (var i = 1; i < rows.length; i++) {
    if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(rows[i][activeIdx])) continue;
    // If no Type column exists, include ALL products (don't filter by type)
    if (typeIdx >= 0) {
      var rowType = String(rows[i][typeIdx] || '').trim();
      if (rowType.toLowerCase() !== type.toLowerCase()) continue;
    }
    var code = codeIdx >= 0 ? String(rows[i][codeIdx] || '').trim() : '';
    var name = nameIdx >= 0 ? String(rows[i][nameIdx] || '').trim() : '';
    if (code) lookup[code] = true;
    if (name) lookup[name] = true;
  }
  Logger.log('SmartPicks _getProductCodesByType(' + type + '): found ' + Object.keys(lookup).length + ' products');
  return lookup;
}

/**
 * Build a mapping of product code → supplier from Products sheet.
 * @param {string} type - 'Sell' or 'Buy'
 * @returns {Object} - { code: supplierName }
 */
function _getProductSupplierMap(type) {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var typeIdx = headers.findIndex(function(h) { return h === 'type' || h === 'product type'; });
  var codeIdx = headers.findIndex(function(h) { return /product.*code|code/i.test(h); });
  var supplierIdx = headers.findIndex(function(h) { return h === 'supplier'; });
  var activeIdx = headers.findIndex(function(h) { return h === 'active'; });

  if (supplierIdx < 0) return {};

  var map = {};
  for (var i = 1; i < rows.length; i++) {
    if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(rows[i][activeIdx])) continue;
    if (typeIdx >= 0) {
      var rowType = String(rows[i][typeIdx] || '').trim();
      if (rowType.toLowerCase() !== type.toLowerCase()) continue;
    }
    var code = codeIdx >= 0 ? String(rows[i][codeIdx] || '').trim() : '';
    var supplier = String(rows[i][supplierIdx] || '').trim();
    if (code && supplier) map[code] = supplier;
  }
  return map;
}

/**
 * Build a mapping of product code → navCategory (NAVCATEGORY || CATEGORY) from Products sheet.
 * @param {string} type - 'Sell' or 'Buy'
 * @returns {Object} - { code: navCategoryName }
 */
function _getProductNavCategoryMap(type) {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var typeIdx = headers.findIndex(function(h) { return h === 'type' || h === 'product type'; });
  var codeIdx = headers.findIndex(function(h) { return /product.*code|code/i.test(h); });
  var navCatIdx = headers.findIndex(function(h) { return /navcategory|nav.*cat/i.test(h); });
  var catIdx = headers.findIndex(function(h) { return h === 'category'; });
  var activeIdx = headers.findIndex(function(h) { return h === 'active'; });

  var map = {};
  for (var i = 1; i < rows.length; i++) {
    if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(rows[i][activeIdx])) continue;
    if (typeIdx >= 0) {
      var rowType = String(rows[i][typeIdx] || '').trim();
      if (rowType.toLowerCase() !== type.toLowerCase()) continue;
    }
    var code = codeIdx >= 0 ? String(rows[i][codeIdx] || '').trim() : '';
    if (!code) continue;
    var navCat = navCatIdx >= 0 ? String(rows[i][navCatIdx] || '').trim() : '';
    if (!navCat && catIdx >= 0) navCat = String(rows[i][catIdx] || '').trim();
    if (navCat) map[code] = navCat;
  }
  return map;
}
function _computeSupplierTopSoldProducts() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var sellProducts = _getProductCodesByType('Sell');
  var supplierMap = _getProductSupplierMap('Sell');
  if (Object.keys(supplierMap).length === 0) return {};

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var tsIdx = headers.indexOf('timestamp');
  var codeIdx = headers.indexOf('product code');
  if (codeIdx === -1) codeIdx = headers.findIndex(function(h) { return /product.*code/i.test(h); });
  var qtyIdx = headers.indexOf('quantity');
  if (qtyIdx === -1) qtyIdx = headers.findIndex(function(h) { return /qty|quantity/i.test(h); });
  var outletIdx = headers.indexOf('outlet');
  var navCatColIdx = headers.findIndex(function(h) { return /navcategory|nav.*cat/i.test(h); });
  var navCatMap = _getProductNavCategoryMap('Sell');

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  // Count per outlet → supplier → navCategory → code
  var byOutletSupplier = {};
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var ts = row[tsIdx] ? new Date(row[tsIdx]) : null;
    if (!ts || ts < cutoff) continue;
    var outlet = outletIdx >= 0 ? String(row[outletIdx] || '').trim() : '';
    var code = codeIdx >= 0 ? String(row[codeIdx] || '').trim() : '';
    if (!code || !sellProducts[code]) continue;
    var supplier = supplierMap[code];
    if (!supplier) continue;
    var qty = Number(row[qtyIdx]) || 1;
    var navCat = (navCatColIdx >= 0 ? String(row[navCatColIdx] || '').trim() : '') || navCatMap[code] || '';

    if (!byOutletSupplier[outlet]) byOutletSupplier[outlet] = {};
    if (!byOutletSupplier[outlet][supplier]) byOutletSupplier[outlet][supplier] = {};
    if (!byOutletSupplier[outlet][supplier][navCat]) byOutletSupplier[outlet][supplier][navCat] = {};
    if (!byOutletSupplier[outlet][supplier][navCat][code]) byOutletSupplier[outlet][supplier][navCat][code] = { count: 0, totalQty: 0 };
    byOutletSupplier[outlet][supplier][navCat][code].count++;
    byOutletSupplier[outlet][supplier][navCat][code].totalQty += qty;
  }

  // For each outlet+supplier+navCategory, rank and take top 10
  var result = {};
  Object.keys(byOutletSupplier).forEach(function(outlet) {
    Object.keys(byOutletSupplier[outlet]).forEach(function(supplier) {
      var key = outlet + '|||' + supplier;
      Object.keys(byOutletSupplier[outlet][supplier]).forEach(function(navCat) {
        var freq = byOutletSupplier[outlet][supplier][navCat];
        var sorted = Object.keys(freq).sort(function(a, b) {
          return freq[b].count - freq[a].count || freq[b].totalQty - freq[a].totalQty;
        });
        var top10 = sorted.slice(0, 10);
        if (top10.length > 0) {
          if (!result[key]) result[key] = {};
          result[key][navCat] = top10;
        }
      });
    });
  });
  return result;
}

/**
 * Compute top 20 sold products per outlet from Orders sheet (last 30 days).
 * Only includes products that exist in Products sheet with Type=Sell.
 */
function _computeTopSoldProducts() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  // Get valid Sell product codes
  var sellProducts = _getProductCodesByType('Sell');

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var tsIdx = headers.indexOf('timestamp');
  var codeIdx = headers.indexOf('product code');
  if (codeIdx === -1) codeIdx = headers.findIndex(function(h) { return /product.*code/i.test(h); });
  var qtyIdx = headers.indexOf('quantity');
  if (qtyIdx === -1) qtyIdx = headers.findIndex(function(h) { return /qty|quantity/i.test(h); });
  var outletIdx = headers.indexOf('outlet');
  var nameIdx = headers.indexOf('product name');
  if (nameIdx === -1) nameIdx = headers.findIndex(function(h) { return /product|name|item/i.test(h); });
  var navCatColIdx = headers.findIndex(function(h) { return /navcategory|nav.*cat/i.test(h); });
  var navCatMap = _getProductNavCategoryMap('Sell');

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  var byOutlet = {};
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var ts = row[tsIdx] ? new Date(row[tsIdx]) : null;
    if (!ts || ts < cutoff) continue;
    var outlet = outletIdx >= 0 ? String(row[outletIdx] || '').trim() : '';
    var code = codeIdx >= 0 ? String(row[codeIdx] || '').trim() : '';
    if (!code && nameIdx >= 0) code = String(row[nameIdx] || '').trim();
    if (!code || !sellProducts[code]) continue;
    var qty = Number(row[qtyIdx]) || 1;
    // Prefer navCategory stored in order row; fall back to Products sheet lookup
    var navCat = (navCatColIdx >= 0 ? String(row[navCatColIdx] || '').trim() : '') || navCatMap[code] || '';

    if (!byOutlet[outlet]) byOutlet[outlet] = {};
    if (!byOutlet[outlet][navCat]) byOutlet[outlet][navCat] = {};
    if (!byOutlet[outlet][navCat][code]) byOutlet[outlet][navCat][code] = { count: 0, totalQty: 0 };
    byOutlet[outlet][navCat][code].count++;
    byOutlet[outlet][navCat][code].totalQty += qty;
  }

  // Return { outlet: { navCategory: [topCodes] } }
  var result = {};
  Object.keys(byOutlet).forEach(function(outlet) {
    result[outlet] = {};
    Object.keys(byOutlet[outlet]).forEach(function(navCat) {
      var freq = byOutlet[outlet][navCat];
      result[outlet][navCat] = Object.keys(freq).sort(function(a, b) {
        return freq[b].count - freq[a].count || freq[b].totalQty - freq[a].totalQty;
      }).slice(0, 20);
    });
  });
  return result;
}

/**
 * Compute top 20 consumed products per outlet from Inventory_Consumption sheet (last 30 days).
 * Only includes products that exist in Products sheet with Type=Buy.
 */
function _computeTopConsumedProducts() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
  if (!sheet || sheet.getLastRow() < 2) return {};

  // Get valid Buy product codes
  var buyProducts = _getProductCodesByType('Buy');

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h) { return String(h).toLowerCase(); });
  var tsIdx = headers.findIndex(function(h) { return /timestamp|date/i.test(h); });
  var codeIdx = headers.findIndex(function(h) { return /product.*code|code/i.test(h); });
  var qtyIdx = headers.findIndex(function(h) { return /quantity/i.test(h); });
  var outletIdx = headers.findIndex(function(h) { return /outlet/i.test(h); });

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  var byOutlet = {};
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var ts = row[tsIdx] ? new Date(row[tsIdx]) : null;
    if (!ts || ts < cutoff) continue;
    var outlet = outletIdx >= 0 ? String(row[outletIdx] || '').trim() : '';
    var code = codeIdx >= 0 ? String(row[codeIdx] || '').trim() : '';
    if (!code) continue;
    // Only count if it's a valid Buy product
    if (!buyProducts[code]) continue;
    var qty = Number(row[qtyIdx]) || 0;
    if (!byOutlet[outlet]) byOutlet[outlet] = {};
    if (!byOutlet[outlet][code]) byOutlet[outlet][code] = { count: 0, totalQty: 0 };
    byOutlet[outlet][code].count++;
    byOutlet[outlet][code].totalQty += qty;
  }

  var result = {};
  Object.keys(byOutlet).forEach(function(outlet) {
    var freq = byOutlet[outlet];
    result[outlet] = Object.keys(freq).sort(function(a, b) {
      return freq[b].count - freq[a].count || freq[b].totalQty - freq[a].totalQty;
    }).slice(0, 20);
  });
  return result;
}

/**
 * Refresh all smart picks data into the SmartPicks sheet.
 * Called by scheduled trigger or admin ad-hoc.
 */
function refreshSmartPicks() {
  var sheet = ensureSmartPicksSheet();

  // Clear existing data (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var newRows = [];

  // Compute orders top picks (now returns { outlet: { navCategory: [codes] } })
  var orderPicks = _computeTopSoldProducts();
  Logger.log('SmartPicks refresh: orderPicks outlets=' + JSON.stringify(Object.keys(orderPicks)));
  Object.keys(orderPicks).forEach(function(outlet) {
    Object.keys(orderPicks[outlet]).forEach(function(navCat) {
      orderPicks[outlet][navCat].forEach(function(code, idx) {
        newRows.push(['orders', outlet, '', navCat, code, idx + 1, now]);
      });
    });
  });

  // Compute per-supplier top picks (now returns { "outlet|||supplier": { navCategory: [codes] } })
  var supplierPicks = _computeSupplierTopSoldProducts();
  var supplierKeys = Object.keys(supplierPicks);
  Logger.log('SmartPicks refresh: supplierPicks keys=' + supplierKeys.length);
  supplierKeys.forEach(function(key) {
    var parts = key.split('|||');
    var outlet = parts[0];
    var supplier = parts[1];
    Object.keys(supplierPicks[key]).forEach(function(navCat) {
      supplierPicks[key][navCat].forEach(function(code, idx) {
        newRows.push(['orders', outlet, supplier, navCat, code, idx + 1, now]);
      });
    });
  });

  // Compute consumption top picks
  var consumptionPicks = _computeTopConsumedProducts();
  Logger.log('SmartPicks refresh: consumptionPicks outlets=' + JSON.stringify(Object.keys(consumptionPicks)));
  Object.keys(consumptionPicks).forEach(function(outlet) {
    consumptionPicks[outlet].forEach(function(code, idx) {
      newRows.push(['consumption', outlet, '', '', code, idx + 1, now]);
    });
  });

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 7).setValues(newRows);
  }

  Logger.log('SmartPicks refreshed: ' + newRows.length + ' rows written at ' + now);
  return { success: true, message: 'Smart picks refreshed', totalRows: newRows.length, updatedAt: now };
}

/**
 * Read cached smart picks for a specific type and outlet.
 * Returns array of product codes.
 */
function getSmartPicks(type, deviceId, username) {
  try {
    var ss = getGlobalSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SMART_PICKS);
    if (!sheet || sheet.getLastRow() < 2) return { codes: [], updatedAt: null };

    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet || '';

    var rows = sheet.getDataRange().getValues();
    // Column layout detection:
    //   7-col: Type, Outlet, Supplier, NavCategory, ProductCode, Rank, UpdatedAt
    //   6-col: Type, Outlet, Supplier, ProductCode, Rank, UpdatedAt  (old)
    //   5-col: Type, Outlet, ProductCode, Rank, UpdatedAt            (very old)
    var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var is7col = headers.length >= 7 && headers[3] === 'navcategory';
    var is6col = !is7col && headers.length >= 6 && headers[2] === 'supplier';
    var codeColIdx    = is7col ? 4 : (is6col ? 3 : 2);
    var rankColIdx    = is7col ? 5 : (is6col ? 4 : 3);
    var updatedAtColIdx = is7col ? 6 : (is6col ? 5 : 4);
    var supplierColIdx = (is7col || is6col) ? 2 : -1;

    var codes = [];
    var updatedAt = null;

    for (var i = 1; i < rows.length; i++) {
      var rowType = String(rows[i][0] || '').trim();
      var rowSupplier = supplierColIdx >= 0 ? String(rows[i][supplierColIdx] || '').trim() : '';
      // For getSmartPicks we only want global rows (no supplier)
      var isGlobal = rowSupplier === '' && rowType === type;
      if (isGlobal && rows[i][1] === userOutlet) {
        codes.push({ code: String(rows[i][codeColIdx] || ''), rank: Number(rows[i][rankColIdx]) || 0 });
        if (!updatedAt) {
          var raw = rows[i][updatedAtColIdx];
          updatedAt = (raw instanceof Date) ? raw.toISOString() : String(raw || '');
        }
      }
    }

    codes.sort(function(a, b) { return a.rank - b.rank; });
    return sanitizeForClient({
      codes: codes.map(function(c) { return c.code; }),
      updatedAt: updatedAt
    });
  } catch (e) {
    Logger.log('getSmartPicks error: ' + e);
    return { codes: [], updatedAt: null };
  }
}

/**
 * Get both orders and consumption smart picks in one call (for login caching).
 * Simplified: reads directly from SmartPicks sheet, matches outlet from session if possible.
 */
function getAllSmartPicks(deviceId, username) {
  try {
    var ss = getGlobalSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SMART_PICKS);
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('getAllSmartPicks: Sheet empty or missing');
      return { orders: [], consumption: [], updatedAt: null };
    }

    // Try to get outlet from session, but don't fail if we can't
    var userOutlet = '';
    try {
      var session = getUserSessionEnhanced(deviceId, username) || {};
      userOutlet = (session.outlet || '').toString().trim().toLowerCase();
    } catch(sessionErr) {
      Logger.log('getAllSmartPicks: session lookup failed: ' + sessionErr);
    }
    Logger.log('getAllSmartPicks: userOutlet="' + userOutlet + '"');

    var rows = sheet.getDataRange().getValues();

    // Column layout detection:
    //   7-col: Type, Outlet, Supplier, NavCategory, ProductCode, Rank, UpdatedAt
    //   6-col: Type, Outlet, Supplier, ProductCode, Rank, UpdatedAt  (old)
    //   5-col: Type, Outlet, ProductCode, Rank, UpdatedAt            (very old)
    var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var is7col = headers.length >= 7 && headers[3] === 'navcategory';
    var is6col = !is7col && headers.length >= 6 && headers[2] === 'supplier';
    var codeColIdx      = is7col ? 4 : (is6col ? 3 : 2);
    var rankColIdx      = is7col ? 5 : (is6col ? 4 : 3);
    var updatedAtColIdx = is7col ? 6 : (is6col ? 5 : 4);
    var supplierColIdx  = (is7col || is6col) ? 2 : -1;
    var navCatColIdx    = is7col ? 3 : -1;

    // Accumulate entries keyed by outlet match
    var orderEntries     = []; // global order picks for user's outlet
    var allOrderEntries  = []; // fallback: all outlets
    var consumptionEntries    = [];
    var allConsumptionEntries = [];
    // supplierNavCat: { supplierName: { navCat: [entries] } }
    var supplierNavCat    = {};
    var allSupplierNavCat = {};
    // navCategoryPicks (global): { navCat: [entries] }
    var navCatEntries    = {};
    var allNavCatEntries = {};

    var updatedAt = null;

    for (var i = 1; i < rows.length; i++) {
      var rowOutlet = String(rows[i][1] || '').trim().toLowerCase();
      var code = String(rows[i][codeColIdx] || '').trim();
      if (!code) continue;
      var rank = Number(rows[i][rankColIdx]) || 0;
      if (!updatedAt) {
        var raw = rows[i][updatedAtColIdx];
        updatedAt = (raw instanceof Date) ? raw.toISOString() : String(raw || '');
      }
      var type = String(rows[i][0] || '').trim();
      var typeLower = type.toLowerCase();

      // Determine supplier
      var supplierName = '';
      if (supplierColIdx >= 0) {
        supplierName = String(rows[i][supplierColIdx] || '').trim();
      } else if (typeLower.indexOf('orders_supplier_') === 0) {
        supplierName = type.substring('orders_supplier_'.length);
        typeLower = 'orders';
      }

      // Determine navCategory
      var navCat = navCatColIdx >= 0 ? String(rows[i][navCatColIdx] || '').trim() : '';

      var entry = { code: code, rank: rank, navCat: navCat };
      var forThisOutlet = userOutlet && rowOutlet === userOutlet;

      if (supplierName) {
        // Supplier-specific pick
        if (forThisOutlet) {
          if (!supplierNavCat[supplierName]) supplierNavCat[supplierName] = {};
          if (!supplierNavCat[supplierName][navCat]) supplierNavCat[supplierName][navCat] = [];
          supplierNavCat[supplierName][navCat].push(entry);
        }
        if (!allSupplierNavCat[supplierName]) allSupplierNavCat[supplierName] = {};
        if (!allSupplierNavCat[supplierName][navCat]) allSupplierNavCat[supplierName][navCat] = [];
        allSupplierNavCat[supplierName][navCat].push(entry);
        continue;
      }

      // Global pick (orders or consumption)
      if (typeLower === 'orders') {
        allOrderEntries.push(entry);
        if (navCat) {
          if (!allNavCatEntries[navCat]) allNavCatEntries[navCat] = [];
          allNavCatEntries[navCat].push(entry);
        }
        if (forThisOutlet) {
          orderEntries.push(entry);
          if (navCat) {
            if (!navCatEntries[navCat]) navCatEntries[navCat] = [];
            navCatEntries[navCat].push(entry);
          }
        }
      } else if (typeLower === 'consumption') {
        allConsumptionEntries.push(entry);
        if (forThisOutlet) consumptionEntries.push(entry);
      }
    }

    // Use outlet-specific if found, otherwise fall back to all
    if (orderEntries.length === 0)      orderEntries      = allOrderEntries;
    if (consumptionEntries.length === 0) consumptionEntries = allConsumptionEntries;
    var finalNavCatEntries    = Object.keys(navCatEntries).length    > 0 ? navCatEntries    : allNavCatEntries;
    var finalSupplierNavCat   = Object.keys(supplierNavCat).length   > 0 ? supplierNavCat   : allSupplierNavCat;

    // Helper to sort + dedupe + extract codes from entries
    function _finalizeCodes(entries) {
      entries.sort(function(a, b) { return a.rank - b.rank; });
      var seen = {};
      return entries.filter(function(e) {
        if (seen[e.code]) return false; seen[e.code] = true; return true;
      }).map(function(e) { return e.code; });
    }

    // Global order codes (flat, for the global ⚡ Smart Picks view)
    var orderCodes      = _finalizeCodes(orderEntries);
    var consumptionCodes = _finalizeCodes(consumptionEntries);

    // navCategoryPicks: { navCat: [codes] }
    var navCategoryPicks = {};
    Object.keys(finalNavCatEntries).forEach(function(nc) {
      navCategoryPicks[nc] = _finalizeCodes(finalNavCatEntries[nc]);
    });

    // supplierPicks: { supplierName: [codes] } — flat merge for backward compat
    // supplierNavCategoryPicks: { supplierName: { navCat: [codes] } }
    var supplierPicks = {};
    var supplierNavCategoryPicks = {};
    Object.keys(finalSupplierNavCat).forEach(function(sup) {
      supplierNavCategoryPicks[sup] = {};
      var allSupCodes = [];
      Object.keys(finalSupplierNavCat[sup]).forEach(function(nc) {
        var codes = _finalizeCodes(finalSupplierNavCat[sup][nc].slice());
        supplierNavCategoryPicks[sup][nc] = codes;
        allSupCodes = allSupCodes.concat(codes);
      });
      // Flat deduped list for backward compat
      var flatSeen = {};
      supplierPicks[sup] = allSupCodes.filter(function(c) {
        if (flatSeen[c]) return false; flatSeen[c] = true; return true;
      });
    });

    var result = {
      orders: orderCodes,
      consumption: consumptionCodes,
      navCategoryPicks: navCategoryPicks,
      supplierPicks: supplierPicks,
      supplierNavCategoryPicks: supplierNavCategoryPicks,
      updatedAt: updatedAt
    };
    Logger.log('getAllSmartPicks: returning ' + result.orders.length + ' orders, ' + result.consumption.length + ' consumption, ' + Object.keys(supplierPicks).length + ' suppliers, ' + Object.keys(navCategoryPicks).length + ' navCats');
    return sanitizeForClient(result);
  } catch (e) {
    Logger.log('getAllSmartPicks error: ' + e);
    return { orders: [], consumption: [], updatedAt: null };
  }
}

// ========== TRIGGER MANAGEMENT ==========

/**
 * Schedule a daily trigger to refresh smart picks.
 * @param {number} hour - Hour to run (0-23), default midnight (0)
 */
function scheduleSmartPicksTrigger(hour) {
  if (hour === undefined || hour === null) hour = 0;
  hour = Math.max(0, Math.min(23, parseInt(hour) || 0));

  // Remove existing smart picks triggers first
  removeSmartPicksTrigger();

  ScriptApp.newTrigger('refreshSmartPicks')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  Logger.log('SmartPicks trigger scheduled at hour ' + hour);
  return { success: true, message: 'Daily trigger set for ' + (hour < 10 ? '0' : '') + hour + ':00' };
}

/**
 * Remove all smart picks triggers.
 */
function removeSmartPicksTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'refreshSmartPicks') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' SmartPicks triggers');
  return { success: true, removed: removed };
}

/**
 * Get current smart picks trigger info.
 */
function getSmartPicksTriggerInfo() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = null;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'refreshSmartPicks') {
      found = {
        id: trigger.getUniqueId(),
        type: trigger.getEventType().toString(),
        source: trigger.getTriggerSource().toString()
      };
    }
  });

  // Get last update time from sheet
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SMART_PICKS);
  var lastUpdated = null;
  var totalRows = 0;
  if (sheet && sheet.getLastRow() > 1) {
    totalRows = sheet.getLastRow() - 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var is7col = headers.length >= 7 && String(headers[3] || '').toLowerCase().trim() === 'navcategory';
    var is6col = !is7col && headers.length >= 6 && String(headers[2] || '').toLowerCase().trim() === 'supplier';
    var updatedAtCol = is7col ? 7 : (is6col ? 6 : 5);
    var lastRow = sheet.getRange(2, updatedAtCol).getValue();
    if (lastRow) lastUpdated = String(lastRow);
  }

  return {
    triggerActive: !!found,
    trigger: found,
    lastUpdated: lastUpdated,
    totalRows: totalRows
  };
}
