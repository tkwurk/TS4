// ========== INVENTORY CONSUMPTION MANAGEMENT ==========

function ensureInventoryConsumptionSheetHeaders() {
  const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
  if (!sheet) {
    throw new Error("Inventory Consumption sheet not found.");
  }
  
  // Check if headers exist
  if (sheet.getLastRow() === 0) {
    // Sheet is empty, add headers
    const headers = ['Timestamp', 'Product Code', 'Quantity Consumed', 'Outlet', 'Username'];
    sheet.appendRow(headers);
    Logger.log("Added headers to Inventory Consumption sheet: " + headers.join(", "));
  } else {
    // Check if headers are correct
    const existingHeaders = sheet.getRange(1, 1, 1, 5).getValues()[0];
    const expectedHeaders = ['Timestamp', 'Product Code', 'Quantity Consumed', 'Outlet', 'Username'];
    
    // If headers don't match, log the discrepancy
    if (JSON.stringify(existingHeaders) !== JSON.stringify(expectedHeaders)) {
      Logger.log("Inventory Consumption sheet headers mismatch:");
      Logger.log("Existing: " + JSON.stringify(existingHeaders));
      Logger.log("Expected: " + JSON.stringify(expectedHeaders));
    }
  }
}

/**
 * Append consumption rows without requiring a live session (scheduler / batch).
 */
function appendConsumptionRows_(items, outlet, username, timestamp) {
  if (!items || !items.length) return 0;
  ensureInventoryConsumptionSheetHeaders();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
  if (!sheet) throw new Error('Inventory Consumption sheet not found');

  var ts = timestamp || new Date();
  var rows = [];
  items.forEach(function(item) {
    var qty = Number(item.quantity) || 0;
    if (qty <= 0) return;
    rows.push([
      formatTimestampForSheet(ts),
      item.productCode || '',
      qty,
      outlet || '',
      username || ''
    ]);
  });
  if (!rows.length) return 0;
  if (typeof appendRowsAtomic_ === 'function') {
    appendRowsAtomic_(sheet, rows);
  } else {
    rows.forEach(function(row) { sheet.appendRow(row); });
  }
  return rows.length;
}

function getBuyTypeProducts() {
  try {
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) {
      throw new Error("Products sheet not found");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }

    const headers = data[0];
    const nameIdx = headers.findIndex(h => /^name$/i.test(h));
    const categoryIdx = headers.findIndex(h => /^category$/i.test(h));
    const navCategoryIdx = headers.findIndex(h => /^navcategory$/i.test(h));
    const typeIdx = headers.findIndex(h => /type/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|^code$/i.test(h));
    const buyUnitIdx = headers.findIndex(h => /^buyunit$/i.test(h));
    const unitIdx = buyUnitIdx >= 0 ? buyUnitIdx : headers.findIndex(h => /^unit$/i.test(h));
    const buyPriceIdx = headers.findIndex(h => /buy.*price|cost/i.test(h));
    const supplierIdx = headers.findIndex(h => /supplier|vendor/i.test(h));
    const activeIdx = headers.findIndex(h => /^active$/i.test(h));

    Logger.log("Product columns - Name: " + nameIdx + ", Category: " + categoryIdx + ", Type: " + typeIdx + ", Product Code: " + productCodeIdx + ", BuyUnit: " + unitIdx + ", Buy Price: " + buyPriceIdx);

    const buyProducts = [];
    
    data.slice(1).forEach((row, index) => {
      if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(row[activeIdx])) return;
      const type = typeIdx >= 0 ? (row[typeIdx] || '').toString().toLowerCase() : '';
      const name = nameIdx >= 0 ? row[nameIdx] : '';
      const category = categoryIdx >= 0 ? row[categoryIdx] : '';
      const navCategory = navCategoryIdx >= 0 ? (row[navCategoryIdx] || '').toString().trim() : '';
      let productCode = productCodeIdx >= 0 ? row[productCodeIdx] : '';
      const unit = unitIdx >= 0 ? row[unitIdx] : '';
      const buyPrice = buyPriceIdx >= 0 ? Number(row[buyPriceIdx]) || 0 : 0;
      const supplier = supplierIdx >= 0 ? (row[supplierIdx] || '').toString().trim() : '';
      if (!productCode && name) {
        productCode = generateProductCode({
          name: name,
          category: category
        });
      }

      // Only include 'buy' type products
      if (type === 'buy' && name && productCode) {
        buyProducts.push({
          productCode: productCode,
          name: name,
          category: category,
          navCategory: navCategory,
          unit: unit,
          buyPrice: buyPrice,
          supplier: supplier
        });
      }
    });

    Logger.log("Found " + buyProducts.length + " buy-type products");
    return buyProducts;

  } catch (error) {
    Logger.log("Error getting buy-type products: " + error.toString());
    throw error;
  }
}

function recordInventoryConsumption(consumptionData, deviceId, username) {
  try {
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      throw new Error("Could not access spreadsheet");
    }

    const sheet = spreadsheet.getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!sheet) {
      throw new Error("Inventory Consumption sheet not found");
    }
    
    // Ensure proper headers
    ensureInventoryConsumptionSheetHeaders();

    if (typeof consumptionData === 'string') {
      consumptionData = JSON.parse(consumptionData);
    }

    const session = getUserSessionEnhanced(deviceId, username) || {};
    const now = new Date();

    Logger.log("Recording inventory consumption - Session: " + JSON.stringify(session));
    Logger.log("User outlet: " + (session.outlet || 'NOT SET'));
    Logger.log("Username: " + (session.username || 'NOT SET'));
    Logger.log("Consumption data: " + JSON.stringify(consumptionData));

    // Record each product consumption
    if (Array.isArray(consumptionData.items)) {
      appendConsumptionRows_(
        consumptionData.items,
        session.outlet || '',
        session.username || '',
        now
      );
    }

    return { success: true, message: "Inventory consumption recorded successfully" };

  } catch (error) {
    Logger.log("Error recording inventory consumption: " + error.toString());
    throw error;
  }
}

function getInventoryConsumptionHistory(deviceId, username) {
  try {
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      throw new Error("Could not access spreadsheet");
    }

    const sheet = spreadsheet.getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!sheet) {
      throw new Error("Inventory Consumption sheet not found");
    }

    let rows = [];
    try {
      rows = sheet.getDataRange().getValues() || [];
    } catch (dataError) {
      Logger.log("Error getting sheet data: " + dataError);
      rows = [];
    }

    if (!Array.isArray(rows) || rows.length < 2) {
      return {
        success: true,
        data: [],
        message: "No inventory consumption history found"
      };
    }

    // Process headers and data
    const headers = (rows[0] || []).map(h => String(h).toLowerCase());
    const timestampIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));
    const quantityIdx = headers.findIndex(h => /quantity/i.test(h));
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));
    const usernameIdx = headers.findIndex(h => /username/i.test(h));

    Logger.log("Inventory consumption history - Headers: " + JSON.stringify(headers));
    Logger.log("Column indices - Timestamp: " + timestampIdx + ", ProductCode: " + productCodeIdx + ", Quantity: " + quantityIdx + ", Outlet: " + outletIdx + ", Username: " + usernameIdx);

    // Get user session for filtering
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;

    // Get product names for display
    const productNames = getProductNamesByCode();

    // Ensure we have valid array data before processing
    const validRows = Array.isArray(rows.slice(1)) ? rows.slice(1) : [];

    // Map data to structured format
    const history = validRows
      .filter(row => Array.isArray(row) && (outletIdx === -1 || row[outletIdx] === userOutlet))
      .map(row => {
        try {
          const productCode = row[productCodeIdx] || "";
          const quantityValue = quantityIdx >= 0 ? row[quantityIdx] : null;
          
          // Debug logging for quantity issues
          Logger.log("Processing row - ProductCode: " + productCode + ", QuantityIdx: " + quantityIdx + ", QuantityValue: '" + quantityValue + "' (type: " + typeof quantityValue + ")");
          
          return {
            timestamp: formatDateForJSON(row[timestampIdx]),
            productCode: productCode,
            productName: productNames[productCode] || productCode,
            quantity: quantityValue !== null && quantityValue !== undefined && quantityValue !== '' ? Number(quantityValue) : 0,
            outlet: row[outletIdx] || "",
            username: row[usernameIdx] || ""
          };
        } catch (rowError) {
          Logger.log("Error processing row: " + rowError + " Row data: " + JSON.stringify(row));
          return null;
        }
      })
      .filter(entry => entry !== null)
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    return {
      success: true,
      data: history,
      count: history.length
    };

  } catch (error) {
    Logger.log("Error in getInventoryConsumptionHistory: " + error);
    return {
      success: false,
      data: [],
      error: error.message
    };
  }
}

function getProductNamesByCode() {
  var productNames = {};
  Repo.readRows('PRODUCTS').forEach(function (row) {
    var name = '';
    var category = '';
    var productCode = '';
    Object.keys(row).forEach(function (key) {
      if (/^(product\s*)?name$/i.test(key.trim())) name = row[key] || name;
      else if (/^category$/i.test(key.trim())) category = row[key] || category;
      else if (/^(product\s*code|code)$/i.test(key.trim())) productCode = row[key] || productCode;
    });

    if (!productCode && name) {
      productCode = generateProductCode({ name: name, category: category });
    }
    if (productCode && name) {
      var codeKey = typeof normalizeProductCodeKey_ === 'function'
        ? normalizeProductCodeKey_(productCode)
        : String(productCode).trim().toUpperCase();
      productNames[codeKey] = name;
    }
  });
  return productNames;
}

function shouldPreferBottomUpOrderScan_(businessDateStr) {
  try {
    var todayStr = formatDateForComparison(new Date());
    if (businessDateStr === todayStr) return true;
    var parts = String(businessDateStr).split('-');
    if (parts.length !== 3) return false;
    var target = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    var today = new Date();
    today.setHours(12, 0, 0, 0);
    target.setHours(12, 0, 0, 0);
    var diffDays = Math.round((today.getTime() - target.getTime()) / 86400000);
    return diffDays >= 0 && diffDays <= 2;
  } catch (e) {
    return false;
  }
}

function buildTodaysConsumptionSummaryFromMap_(consumptionMap, productNames, productCategories) {
  var items = [];
  var categorySummary = {};
  var totalQuantity = 0;
  var totalItemsCount = 0;

  Object.keys(consumptionMap || {}).forEach(function(productCode) {
    var quantity = Number(consumptionMap[productCode]) || 0;
    if (!productCode || quantity <= 0) return;

    items.push({
      productCode: productCode,
      productName: productNames[productCode] || productCode,
      quantity: quantity
    });

    var category = productCategories[productCode] || 'Uncategorized';
    if (!categorySummary[category]) {
      categorySummary[category] = {
        category: category,
        totalQuantity: 0,
        itemCount: 0,
        items: {}
      };
    }
    if (!categorySummary[category].items[productCode]) {
      categorySummary[category].items[productCode] = {
        productCode: productCode,
        productName: productNames[productCode] || productCode,
        quantity: 0
      };
      categorySummary[category].itemCount++;
      totalItemsCount++;
    }
    categorySummary[category].items[productCode].quantity += quantity;
    categorySummary[category].totalQuantity += quantity;
    totalQuantity += quantity;
  });

  Object.keys(categorySummary).forEach(function(category) {
    categorySummary[category].items = Object.values(categorySummary[category].items);
  });

  return {
    totalItems: totalItemsCount,
    totalQuantity: totalQuantity,
    items: items,
    categories: categorySummary
  };
}

function getTodaysConsumptionSummary(deviceId, username) {
  try {
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;

    const todayStr = formatDateForComparison(new Date());
    const productNames = typeof getCachedProductNamesByCode_ === 'function'
      ? getCachedProductNamesByCode_()
      : getProductNamesByCode();

    if (typeof readInventoryConsumptionForDate_ === 'function') {
      const scanned = readInventoryConsumptionForDate_(todayStr, userOutlet);
      const built = buildTodaysConsumptionSummaryFromMap_(
        scanned.consumptionMap,
        productNames,
        {}
      );
      return {
        totalItems: built.totalItems,
        totalQuantity: built.totalQuantity,
        items: built.items
      };
    }

    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!sheet) {
      return { totalItems: 0, totalQuantity: 0, items: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { totalItems: 0, totalQuantity: 0, items: [] };
    }

    const headers = data[0];
    const timestampIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));
    const quantityIdx = headers.findIndex(h => /quantity/i.test(h));
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));

    const consumptionMap = {};
    let totalQuantity = 0;

    data.slice(1).forEach(row => {
      // Filter by outlet
      if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;

      const timestamp = row[timestampIdx];
      if (!timestamp) return;

      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
      if (isNaN(date.getTime())) return;

      const dateStr = formatDateForComparison(date);
      if (dateStr !== todayStr) return;

      const productCode = row[productCodeIdx] || "";
      const quantity = Number(row[quantityIdx]) || 0;

      if (productCode && quantity > 0) {
        if (!consumptionMap[productCode]) {
          consumptionMap[productCode] = {
            productCode: productCode,
            productName: productNames[productCode] || productCode,
            quantity: 0
          };
        }
        consumptionMap[productCode].quantity += quantity;
        totalQuantity += quantity;
      }
    });

    const items = Object.values(consumptionMap);

    return {
      totalItems: items.length,
      totalQuantity: totalQuantity,
      items: items
    };

  } catch (error) {
    Logger.log("Error getting today's consumption summary: " + error.toString());
    return { totalItems: 0, totalQuantity: 0, items: [] };
  }
}

/**
 * Keyfigure columns for Daily Closing — reuses today's scans; avoids a second Orders BOM pass.
 */
function buildDailyClosingKeyfigures_(outlet, options) {
  options = options || {};
  var buyProducts = options.buyProducts || [];
  var todayPurchased = options.todayPurchased || {};
  var todayConsumed = options.todayConsumed || {};
  var todaySold = options.todaySold || {};
  var yesterdayStr = options.yesterdayStr || (typeof getYesterdayBusinessDateStr_ === 'function'
    ? getYesterdayBusinessDateStr_()
    : formatDateForComparison(new Date()));
  var snapshot = options.snapshot || (typeof readStockHistorySnapshot_ === 'function'
    ? readStockHistorySnapshot_(outlet)
    : { asOfDate: null, byCode: {} });

  var purchaseQty = {};
  var consumedMap = {};
  var soldMap = {};
  var needsBootstrap = false;
  var stockSource = 'History+Today';
  var historyAsOfDate = snapshot.asOfDate || null;

  if (!snapshot.asOfDate || snapshot.asOfDate < yesterdayStr) {
    if (typeof getStockTotalsDirectSheet_ === 'function') {
      var direct = getStockTotalsDirectSheet_(outlet);
      purchaseQty = direct.purchased || {};
      consumedMap = direct.consumed || {};
      soldMap = direct.sold || {};
      needsBootstrap = true;
      stockSource = direct.stockSource || 'DirectSheet';
    } else {
      purchaseQty = typeof _readPurchaseQtyByCode_ === 'function' ? _readPurchaseQtyByCode_(outlet) : {};
      consumedMap = typeof _readRecordedConsumptionQtyByCode_ === 'function'
        ? _readRecordedConsumptionQtyByCode_(outlet) : {};
      soldMap = {};
      needsBootstrap = true;
      stockSource = 'DirectSheet';
    }
    // Stock_History is not yet current, so the direct/fallback path carries no
    // sold figures. purchaseQty/consumedMap are already lifetime (they read the
    // full sheets), so only merge today's sold to reflect the current day.
    Object.keys(todaySold || {}).forEach(function(k) {
      var nk = typeof normalizeProductCodeKey_ === 'function' ? normalizeProductCodeKey_(k) : String(k || '').trim();
      if (!nk) return;
      soldMap[nk] = Math.round(((Number(soldMap[nk]) || 0) + (Number(todaySold[k]) || 0)) * 1000) / 1000;
    });
  } else {
    var codes = {};
    Object.keys(snapshot.byCode || {}).forEach(function(k) { codes[k] = true; });
    Object.keys(todayPurchased).forEach(function(k) { codes[k] = true; });
    Object.keys(todayConsumed).forEach(function(k) { codes[k] = true; });
    Object.keys(todaySold).forEach(function(k) { codes[k] = true; });

    Object.keys(codes).forEach(function(code) {
      var hist = (snapshot.byCode && snapshot.byCode[code]) || {};
      var round = function(n) { return Math.round((Number(n) || 0) * 1000) / 1000; };
      purchaseQty[code] = round((hist.totalPurchased || 0) + (Number(todayPurchased[code]) || 0));
      consumedMap[code] = round((hist.totalConsumed || 0) + (Number(todayConsumed[code]) || 0));
      soldMap[code] = round((hist.totalSold || 0) + (Number(todaySold[code]) || 0));
    });
  }

  var reorderLookup = typeof _readKeyfigureReorderByCode_ === 'function'
    ? _readKeyfigureReorderByCode_(outlet) : {};
  var recipeItemCodes = {};
  if (!needsBootstrap && typeof buildRecipeItemCodeSet_ === 'function') {
    recipeItemCodes = buildRecipeItemCodeSet_(
      typeof getCachedBomRecipes_ === 'function' ? getCachedBomRecipes_() : { lines: [] },
      { lines: [] }
    );
  }

  var items = [];
  buyProducts.forEach(function(p) {
    var code = (p.productCode || '').toString().trim();
    if (!code) return;
    var nk = typeof normalizeProductCodeKey_ === 'function' ? normalizeProductCodeKey_(code) : code;
    var totalPurchased = Math.round((purchaseQty[nk] || 0) * 1000) / 1000;
    var totalSold = Math.round((soldMap[nk] || 0) * 1000) / 1000;
    var totalConsumed = Math.round((consumedMap[nk] || 0) * 1000) / 1000;
    items.push({
      name: p.name || code,
      productCode: code,
      category: p.category || '',
      navCategory: p.navCategory || '',
      unit: p.unit || '',
      totalPurchased: totalPurchased,
      totalConsumed: totalConsumed,
      totalSold: totalSold,
      remaining: Math.round((totalPurchased - totalConsumed) * 1000) / 1000,
      reorderQty: reorderLookup[nk] || 0,
      supplier: (p.supplier || '').toString().trim(),
      hasRecipe: !!recipeItemCodes[nk]
    });
  });

  items.sort(function(a, b) {
    return (a.remaining || 0) - (b.remaining || 0);
  });

  return {
    items: items,
    stockSource: stockSource,
    historyAsOfDate: historyAsOfDate,
    needsBootstrap: needsBootstrap
  };
}

/**
 * Single server call for Daily Closing — one Orders scan, one Consumption scan.
 */
function getDailyClosingBundle(deviceId, username) {
  try {
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet || '';
    if (typeof resolveUserOutlet_ === 'function' && session.username) {
      userOutlet = resolveUserOutlet_(session.username, userOutlet);
    }

    var todayStr = formatDateForComparison(new Date());
    var yesterdayStr = typeof getYesterdayBusinessDateStr_ === 'function'
      ? getYesterdayBusinessDateStr_()
      : todayStr;

    // Auto-refresh Stock_History through yesterday on navigation to Daily Closing.
    if (typeof ensureStockHistoryCurrentForOutlet_ === 'function') {
      ensureStockHistoryCurrentForOutlet_(userOutlet);
    }

    var buyProducts = typeof getCachedBuyTypeProducts_ === 'function'
      ? getCachedBuyTypeProducts_()
      : (typeof getBuyTypeProducts === 'function' ? getBuyTypeProducts() : []);
    var productNames = typeof getCachedProductNamesByCode_ === 'function'
      ? getCachedProductNamesByCode_()
      : (typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {});
    var bomRecipes = typeof getCachedBomRecipes_ === 'function'
      ? getCachedBomRecipes_()
      : (typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : { byUnitProduct: {} });
    var usageMatrix = typeof getCachedBomUsageMatrix_ === 'function'
      ? getCachedBomUsageMatrix_()
      : { byBaseAndFlavor: {}, flavorCodes: {}, lines: [] };

    var buyByCode = {};
    buyProducts.forEach(function(p) {
      var k = typeof normalizeProductCodeKey_ === 'function'
        ? normalizeProductCodeKey_(p.productCode)
        : (p.productCode || '');
      if (k) buyByCode[k] = p;
    });

    var consumptionScanned = typeof readInventoryConsumptionForDate_ === 'function'
      ? readInventoryConsumptionForDate_(todayStr, userOutlet)
      : { consumptionMap: {} };
    var summaryBuilt = buildTodaysConsumptionSummaryFromMap_(
      consumptionScanned.consumptionMap || {},
      productNames,
      {}
    );

    var preferBottomUp = typeof shouldPreferBottomUpOrderScan_ === 'function'
      ? shouldPreferBottomUpOrderScan_(todayStr)
      : true;
    var combined = typeof computeCombinedConsumptionForOutlet_ === 'function'
      ? computeCombinedConsumptionForOutlet_(
          todayStr,
          userOutlet,
          bomRecipes,
          usageMatrix,
          productNames,
          new Date(),
          {
            productIndex: typeof getCachedProductCodeIndex_ === 'function'
              ? getCachedProductCodeIndex_()
              : null,
            preferBottomUp: preferBottomUp
          }
        )
      : { consumptionMap: {}, details: [] };

    var prefillItems = typeof buildConsumptionPrefillLinesFromDetails_ === 'function'
      ? buildConsumptionPrefillLinesFromDetails_(combined.details || [], buyByCode, productNames)
      : [];

    var snapshot = typeof readStockHistorySnapshot_ === 'function'
      ? readStockHistorySnapshot_(userOutlet)
      : { asOfDate: null, byCode: {} };
    var todayPurchased = {};
    if (snapshot.asOfDate && snapshot.asOfDate >= yesterdayStr) {
      var purchaseIndex = typeof buildPurchaseDateIndexForOutlet_ === 'function'
        ? buildPurchaseDateIndexForOutlet_(userOutlet)
        : {};
      todayPurchased = purchaseIndex[todayStr] || {};
    }

    var todayConsumed = {};
    Object.keys(consumptionScanned.consumptionMap || {}).forEach(function(k) {
      var nk = typeof normalizeProductCodeKey_ === 'function'
        ? normalizeProductCodeKey_(k)
        : String(k || '').trim();
      if (!nk) return;
      todayConsumed[nk] = (todayConsumed[nk] || 0) + (Number(consumptionScanned.consumptionMap[k]) || 0);
    });

    var todaySold = combined.consumptionMap || {};
    var keyfigures = buildDailyClosingKeyfigures_(userOutlet, {
      buyProducts: buyProducts,
      snapshot: snapshot,
      yesterdayStr: yesterdayStr,
      todayPurchased: todayPurchased,
      todayConsumed: todayConsumed,
      todaySold: todaySold
    });

    var smartPicks = typeof getAllSmartPicks === 'function'
      ? getAllSmartPicks(deviceId, username)
      : { orders: [], consumption: [] };

    return {
      success: true,
      products: buyProducts,
      autoPrefill: {
        success: true,
        date: todayStr,
        totalItems: prefillItems.length,
        items: prefillItems
      },
      summary: {
        totalItems: summaryBuilt.totalItems,
        totalQuantity: summaryBuilt.totalQuantity,
        items: summaryBuilt.items
      },
      keyfigures: keyfigures,
      smartPicks: smartPicks
    };
  } catch (error) {
    Logger.log('getDailyClosingBundle error: ' + error.toString());
    return {
      success: false,
      error: error.message || error.toString(),
      products: [],
      autoPrefill: { success: false, date: '', totalItems: 0, items: [] },
      summary: { totalItems: 0, totalQuantity: 0, items: [] },
      keyfigures: { items: [] },
      smartPicks: { orders: [], consumption: [] }
    };
  }
}

function getDailyClosingBundleSecure(params) {
  return withAgentGuard_('report.dailyClosing', params, function () {
    var deviceId = params && (params.deviceId || (params._authContext && params._authContext.deviceId));
    var username = params && (params.username || (params._authContext && params._authContext.username));
    // Retry only the idempotent read; reflection flags any negative stock figures.
    return ErrorHandler.retry(function () {
      return getDailyClosingBundle(deviceId, username);
    }, 3);
  }, {
    reflect: function (result) { return ReflectionAgent.validateStock(result); }
  });
}

function getTodaysAutoConsumptionPrefill(deviceId, username) {
  try {
    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet || '';
    if (typeof resolveUserOutlet_ === 'function' && session.username) {
      userOutlet = resolveUserOutlet_(session.username, userOutlet);
    }

    var todayStr = formatDateForComparison(new Date());
    var items = typeof buildConsumptionPrefillItems_ === 'function'
      ? buildConsumptionPrefillItems_(todayStr, userOutlet, {
          preferBottomUp: shouldPreferBottomUpOrderScan_(todayStr)
        })
      : [];

    return {
      success: true,
      date: todayStr,
      totalItems: items.length,
      items: items
    };
  } catch (error) {
    Logger.log('getTodaysAutoConsumptionPrefill error: ' + error.toString());
    return { success: false, date: '', totalItems: 0, items: [], error: error.message || error.toString() };
  }
}

function getTodaysConsumptionSummaryByCategory(deviceId, username) {
  try {
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;

    const todayStr = formatDateForComparison(new Date());
    const productNames = typeof getCachedProductNamesByCode_ === 'function'
      ? getCachedProductNamesByCode_()
      : getProductNamesByCode();
    const productCategories = getProductCategoriesByCode();

    if (typeof readInventoryConsumptionForDate_ === 'function') {
      const scanned = readInventoryConsumptionForDate_(todayStr, userOutlet);
      const built = buildTodaysConsumptionSummaryFromMap_(
        scanned.consumptionMap,
        productNames,
        productCategories
      );
      return {
        totalItems: built.totalItems,
        totalQuantity: built.totalQuantity,
        categories: built.categories
      };
    }

    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!sheet) {
      return { totalItems: 0, totalQuantity: 0, categories: {} };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { totalItems: 0, totalQuantity: 0, categories: {} };
    }

    const headers = data[0];
    const timestampIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));
    const quantityIdx = headers.findIndex(h => /quantity/i.test(h));
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));

    const categorySummary = {};
    let totalQuantity = 0;
    let totalItemsCount = 0;

    data.slice(1).forEach(row => {
      // Filter by outlet
      if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) return;

      const timestamp = row[timestampIdx];
      if (!timestamp) return;

      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
      if (isNaN(date.getTime())) return;

      const dateStr = formatDateForComparison(date);
      if (dateStr !== todayStr) return;

      const productCode = row[productCodeIdx] || "";
      const quantity = Number(row[quantityIdx]) || 0;

      if (productCode && quantity > 0) {
        const category = productCategories[productCode] || "Uncategorized";
        
        if (!categorySummary[category]) {
          categorySummary[category] = {
            category: category,
            totalQuantity: 0,
            itemCount: 0,
            items: {}
          };
        }

        if (!categorySummary[category].items[productCode]) {
          categorySummary[category].items[productCode] = {
            productCode: productCode,
            productName: productNames[productCode] || productCode,
            quantity: 0
          };
          categorySummary[category].itemCount++;
          totalItemsCount++;
        }

        categorySummary[category].items[productCode].quantity += quantity;
        categorySummary[category].totalQuantity += quantity;
        totalQuantity += quantity;
      }
    });

    // Convert items to arrays
    Object.keys(categorySummary).forEach(category => {
      categorySummary[category].items = Object.values(categorySummary[category].items);
    });

    return {
      totalItems: totalItemsCount,
      totalQuantity: totalQuantity,
      categories: categorySummary
    };

  } catch (error) {
    Logger.log("Error getting today's consumption summary by category: " + error.toString());
    return { totalItems: 0, totalQuantity: 0, categories: {} };
  }
}

function getProductCategoriesByCode() {
  try {
    const sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) {
      return {};
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {};
    }

    const headers = data[0];
    const nameIdx = headers.findIndex(h => /name/i.test(h));
    const categoryIdx = headers.findIndex(h => /category/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));

    const productCategories = {};
    
    data.slice(1).forEach(row => {
      const name = nameIdx >= 0 ? row[nameIdx] : '';
      const category = categoryIdx >= 0 ? row[categoryIdx] : '';
      let productCode = productCodeIdx >= 0 ? row[productCodeIdx] : '';

      // Generate product code if not available
      if (!productCode && name) {
        productCode = generateProductCode({
          name: name,
          category: category
        });
      }

      if (productCode) {
        productCategories[productCode] = category || "Uncategorized";
      }
    });

    return productCategories;

  } catch (error) {
    Logger.log("Error getting product categories: " + error.toString());
    return {};
  }
}

/**
 * SecurityWrapper-compatible inventory consumption recording function
 * @param {Object} params - Parameters object from SecurityWrapper (contains _authContext and consumption data)
 * @returns {Object} Success response or throws error
 */
function recordInventoryConsumptionSecure(params) {
  try {
    Logger.log("recordInventoryConsumptionSecure called with params: " + JSON.stringify(params));
    
    // Extract the consumption data and auth context from SecurityWrapper enhanced parameters
    var consumptionData = params;
    var authContext = params._authContext;

    var idemKey = typeof buildIdempotencyCacheKey_ === 'function'
      ? buildIdempotencyCacheKey_('inv_consume_', params.clientRequestId)
      : '';
    if (idemKey && typeof getIdempotentResult_ === 'function') {
      var cachedConsume = getIdempotentResult_(idemKey);
      if (cachedConsume) {
        cachedConsume.duplicate = true;
        return cachedConsume;
      }
    }
    
    Logger.log("Auth context: " + JSON.stringify(authContext));
    Logger.log("Consumption data: " + JSON.stringify(consumptionData));
    
    // Validate authentication context
    if (!authContext || !authContext.username || !authContext.outlet) {
      throw new Error("Authentication required: Missing username or outlet information");
    }
    
    // Set up session context for the original function to use
    const originalGetUserSession = getUserSessionEnhanced;
    getUserSessionEnhanced = function() {
      return {
        username: authContext.username,
        outlet: authContext.outlet,
        deviceId: authContext.deviceId
      };
    };
    
    try {
      // Call the original recordInventoryConsumption function
      const result = recordInventoryConsumption(consumptionData);
      
      Logger.log("recordInventoryConsumptionSecure completed successfully for user: " + authContext.username);
      var response = { 
        success: true, 
        message: "Inventory consumption recorded successfully",
        username: authContext.username,
        outlet: authContext.outlet,
        ...result
      };
      if (idemKey && typeof putIdempotentResult_ === 'function') {
        putIdempotentResult_(idemKey, response, 600);
      }
      return response;
    } finally {
      // Restore original getUserSessionEnhanced function
      getUserSessionEnhanced = originalGetUserSession;
    }
    
  } catch (error) {
    Logger.log("recordInventoryConsumptionSecure error: " + error.toString());
    throw error;
  }
}

/**
 * Returns the top 15 most frequently consumed products in the last 30 days for the user's outlet.
 */
function getTopConsumedProducts(deviceId, username) {
  try {
    var spreadsheet = getGlobalSpreadsheet();
    var sheet = spreadsheet.getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var session = getUserSessionEnhanced(deviceId, username) || {};
    var userOutlet = session.outlet;
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0].map(function(h) { return String(h).toLowerCase(); });
    var tsIdx = headers.findIndex(function(h) { return /timestamp|date/i.test(h); });
    var codeIdx = headers.findIndex(function(h) { return /product.*code|code/i.test(h); });
    var qtyIdx = headers.findIndex(function(h) { return /quantity/i.test(h); });
    var outletIdx = headers.findIndex(function(h) { return /outlet/i.test(h); });

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);

    var freq = {};
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (outletIdx >= 0 && userOutlet && row[outletIdx] !== userOutlet) continue;
      var ts = row[tsIdx] ? new Date(row[tsIdx]) : null;
      if (!ts || ts < cutoff) continue;
      var code = String(row[codeIdx] || '');
      if (!code) continue;
      var qty = Number(row[qtyIdx]) || 0;
      if (!freq[code]) freq[code] = { count: 0, totalQty: 0 };
      freq[code].count++;
      freq[code].totalQty += qty;
    }

    var sorted = Object.keys(freq).sort(function(a, b) {
      return freq[b].count - freq[a].count || freq[b].totalQty - freq[a].totalQty;
    }).slice(0, 15);

    return sorted;
  } catch (e) {
    Logger.log('getTopConsumedProducts error: ' + e);
    return [];
  }
}
