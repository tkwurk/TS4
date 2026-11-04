function getUsageReportData(deviceId, username) {
  try {
    Logger.log("=== Getting Usage Report Data (Sales + Purchases) ===");
    
    const ordersSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ORDERS);
    const purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    
    if (!ordersSheet) {
      Logger.log("Orders sheet not found");
      return [];
    }

    // Get current user's outlet
    const session = getUserSessionEnhanced(deviceId, username) || {};
    const userOutlet = session.outlet;
    
    const data = ordersSheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    const headers = data.shift();

    // Find necessary columns
    const dateIdx = headers.findIndex(h => /timestamp|date|tanggal|waktu|time/i.test(h));
    const catIdx = headers.findIndex(h => /category|kategori/i.test(h));
    const nameIdx = headers.findIndex(h => /name|product|item/i.test(h));
    const qtyIdx = headers.findIndex(h => /qty|quantity|jumlah/i.test(h));
    const priceIdx = headers.findIndex(h => /^price$/i.test(h) || /unit.*price|sell.*price/i.test(h));
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));
    const productCodeIdx = headers.findIndex(h => /product.*code|code/i.test(h));

    // Aggregate usage data by month and product
    const usageMap = {};

    data.forEach(function(row) {
      // Filter by outlet if user has one set
      if (outletIdx !== -1 && userOutlet && row[outletIdx] !== userOutlet) {
        return;
      }

      const d = row[dateIdx];
      const category = row[catIdx];
      const name = row[nameIdx];
      const qty = Number(row[qtyIdx]) || 0;
      const price = priceIdx >= 0 ? Number(row[priceIdx]) || 0 : 0;
      let productCode = productCodeIdx >= 0 ? row[productCodeIdx] : '';
      
      // Generate product code if not available in the order data
      if (!productCode && (category || name)) {
        productCode = generateProductCode({
          name: name,
          category: category
        });
      }
      
      if (!d || !category || !name || qty === 0 || !productCode) {
        return;
      }
      
      const dateObj = (d instanceof Date) ? d : new Date(d);
      if (isNaN(dateObj.getTime())) {
        return;
      }
      
      const month = dateObj.getFullYear() + '-' + ('0' + (dateObj.getMonth() + 1)).slice(-2);
      
      // Group by month, product code, AND price bucket (zero vs non-zero)
      const priceKey = price > 0 ? 'paid' : 'free';
      const key = month + '||' + productCode + '||' + priceKey;
      
      if (!usageMap[key]) {
        usageMap[key] = {
          month: month,
          product: name,
          productCode: productCode,
          category: category,
          isZeroPrice: price === 0,
          qty: 0
        };
      }
      
      usageMap[key].qty += qty;
    });

    // Process purchase data if purchases sheet exists (chicken tracking)
    if (purchasesSheet) {
      const purchaseData = purchasesSheet.getDataRange().getValues();
      if (purchaseData.length > 1) {
        const purchaseHeaders = purchaseData.shift();
        const pDateIdx = purchaseHeaders.findIndex(h => /timestamp|date/i.test(h));
        const pProductIdx = purchaseHeaders.findIndex(h => /product|item/i.test(h));
        const pQtyIdx = purchaseHeaders.findIndex(h => /qty|quantity/i.test(h));
        const pOutletIdx = purchaseHeaders.findIndex(h => /outlet/i.test(h));
        
        purchaseData.forEach(function(row, rowIndex) {
          // Filter by outlet if user has one set
          if (pOutletIdx !== -1 && userOutlet && row[pOutletIdx] !== userOutlet) {
            return;
          }

          const d = row[pDateIdx];
          const productName = row[pProductIdx];
          const qty = Number(row[pQtyIdx]) || 0;
          
          if (!d || !productName || qty === 0) {
            return;
          }
          
          const dateObj = (d instanceof Date) ? d : new Date(d);
          if (isNaN(dateObj.getTime())) {
            return;
          }
          
          const month = dateObj.getFullYear() + '-' + ('0' + (dateObj.getMonth() + 1)).slice(-2);
          
          // Only track chicken purchases (AYAM HOT EKOR and AYAM ORI EKOR)
          const upperProductName = productName.toUpperCase();
          if (upperProductName.includes('AYAM HOT EKOR') || upperProductName.includes('AYAM ORI EKOR')) {
            
            // Create a key for grouping by month and product
            const key = month + '||' + upperProductName;
            
            if (!usageMap[key]) {
              usageMap[key] = {
                month: month,
                product: productName,
                productCode: upperProductName.replace(/\s+/g, '_'),
                category: 'PURCHASES',
                qty: 0
              };
            }
            
            // For purchases, we add to qty (these are chickens purchased)
            usageMap[key].qty += qty;
          }
        });
      }
    }

    const result = Object.values(usageMap);
    Logger.log("Returning " + result.length + " usage report rows (including purchases)");
    
    return result;
    
  } catch (error) {
    Logger.log("Error getting usage report data: " + error.toString());
    return [];
  }
  
}
  // Test/Debug function for usage report
function testUsageReport() {
  Logger.log("=== Testing Usage Report ===");
  
  try {
    var result = getUsageReportData();
    Logger.log("Usage report test result count: " + result.length);
    Logger.log("Usage report test result: " + JSON.stringify(result.slice(0, 3))); // Show first 3 items
    return result;
  } catch (error) {
    Logger.log("Error in usage report test: " + error.toString());
    return [];
  }
}

/**
 * Quick Ayam Stock — BOM-based remaining stock for buy-type products.
 */
function _buildProductSupplierLookup_() {
  var lookup = { byCode: {}, byName: {} };
  try {
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return lookup;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return lookup;
    var headers = data[0].map(function(h) { return (h || '').toString().trim(); });
    var codeIdx = headers.findIndex(function(h) { return /product.*code|^code$/i.test(h); });
    var nameIdx = headers.findIndex(function(h) { return /^name$/i.test(h); });
    var supIdx = headers.findIndex(function(h) { return /supplier|vendor/i.test(h); });
    if (supIdx < 0) return lookup;
    for (var i = 1; i < data.length; i++) {
      var sup = (data[i][supIdx] || '').toString().trim();
      if (!sup) continue;
      if (codeIdx >= 0) {
        var code = (data[i][codeIdx] || '').toString().trim();
        if (code) lookup.byCode[code] = sup;
      }
      if (nameIdx >= 0) {
        var name = (data[i][nameIdx] || '').toString().trim();
        if (name) lookup.byName[name.toLowerCase()] = sup;
      }
    }
  } catch (e) {
    Logger.log('_buildProductSupplierLookup_ error: ' + e.toString());
  }
  return lookup;
}

/**
 * Sum purchase qty by product code for one outlet.
 */
function _readPurchaseQtyByCode_(outlet) {
  var purchaseQty = {};
  try {
    var purchasesSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PURCHASES);
    if (!purchasesSheet || purchasesSheet.getLastRow() < 2) return purchaseQty;
    var purchaseData = purchasesSheet.getDataRange().getValues();
    var phHeaders = purchaseData.shift();
    var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
    var phQtyIdx = phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
    var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
    purchaseData.forEach(function(row) {
      if (phOutletIdx !== -1 && outlet) {
        if (typeof outletRowMatchesUser_ === 'function') {
          if (!outletRowMatchesUser_(row[phOutletIdx], outlet)) return;
        } else if (row[phOutletIdx] !== outlet) return;
      }
      var code = phCodeIdx >= 0 ? (row[phCodeIdx] || '').toString().trim() : '';
      var qty = Number(row[phQtyIdx]) || 0;
      if (!code || qty === 0) return;
      var nk = typeof normalizeProductCodeKey_ === 'function' ? normalizeProductCodeKey_(code) : code;
      purchaseQty[nk] = (purchaseQty[nk] || 0) + qty;
    });
  } catch (e) {
    Logger.log('_readPurchaseQtyByCode_ error: ' + e.toString());
  }
  return purchaseQty;
}

/**
 * Reorder qty from Keyfigures column J — no flush (flush recalculates entire workbook).
 */
function _readKeyfigureReorderByCode_(outlet) {
  var lookup = {};
  try {
    var sheet = getGlobalSpreadsheet().getSheetByName('Keyfigures');
    if (!sheet || sheet.getLastRow() < 2) return lookup;
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    for (var i = 0; i < data.length; i++) {
      var productCode = (data[i][1] || '').toString().trim();
      if (!productCode) continue;
      var nk = typeof normalizeProductCodeKey_ === 'function' ? normalizeProductCodeKey_(productCode) : productCode;
      lookup[nk] = parseFloat((data[i][9] || '').toString().trim()) || 0;
    }
  } catch (e) {
    Logger.log('_readKeyfigureReorderByCode_ error: ' + e.toString());
  }
  return lookup;
}

/**
 * Sum recorded inventory consumption by product code for one outlet (lifetime).
 */
function _readRecordedConsumptionQtyByCode_(outlet) {
  var consumptionQty = {};
  try {
    var consumptionSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
    if (!consumptionSheet || consumptionSheet.getLastRow() < 2) return consumptionQty;
    var consData = consumptionSheet.getDataRange().getValues();
    var ch = consData[0];
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
      var cCode = cCodeIdx >= 0 ? (crow[cCodeIdx] || '').toString().trim() : '';
      var cQty = cQtyIdx >= 0 ? (Number(crow[cQtyIdx]) || 0) : 0;
      if (!cCode || cQty === 0) continue;
      var nk = typeof normalizeProductCodeKey_ === 'function' ? normalizeProductCodeKey_(cCode) : cCode;
      consumptionQty[nk] = (consumptionQty[nk] || 0) + cQty;
    }
  } catch (e) {
    Logger.log('_readRecordedConsumptionQtyByCode_ error: ' + e.toString());
  }
  return consumptionQty;
}

/**
 * Quick stock report — BOM-based remaining stock for buy-type products.
 * totalSold = BOM consumption from orders; totalConsumed = Inventory_Consumption sheet.
 * remaining = totalPurchased - totalConsumed.
 */
function getAyamStockQuick(params) {
  try {
    var userOutlet = '';
    if (params && typeof params === 'object') {
      var deviceId = params.deviceId || (params._authContext && params._authContext.deviceId) || (params._validatedUser && params._validatedUser.deviceId) || null;
      var resolvedUsername = params.username || (params._authContext && params._authContext.username) || (params._validatedUser && params._validatedUser.username) || null;
      var session = getUserSessionEnhanced(deviceId, resolvedUsername) || {};
      userOutlet = session.outlet || '';
      if (typeof resolveUserOutlet_ === 'function' && session.username) {
        userOutlet = resolveUserOutlet_(session.username, userOutlet);
      }
    }

    // Auto-refresh Stock_History through yesterday on navigation to the stock
    // report so the snapshot is current and today's activity is added live.
    if (typeof ensureStockHistoryCurrentForOutlet_ === 'function') {
      ensureStockHistoryCurrentForOutlet_(userOutlet);
    }

    var buyProducts = typeof getCachedBuyTypeProducts_ === 'function'
      ? getCachedBuyTypeProducts_()
      : (typeof getBuyTypeProducts === 'function' ? getBuyTypeProducts() : []);
    if (!buyProducts.length) {
      return { items: [], stockSource: 'BOM', error: 'No buy-type products found' };
    }

    var purchaseQty = {};
    var salesConsumptionMap = {};
    var recordedConsumptionMap = {};
    var stockSource = 'History+Today';
    var historyAsOfDate = null;
    var needsBootstrap = false;

    if (typeof getCachedStockTotalsForOutlet_ === 'function') {
      var stockTotals = getCachedStockTotalsForOutlet_(userOutlet);
      purchaseQty = stockTotals.purchased || {};
      recordedConsumptionMap = stockTotals.consumed || {};
      salesConsumptionMap = stockTotals.sold || {};
      stockSource = stockTotals.stockSource || stockSource;
      historyAsOfDate = stockTotals.historyAsOfDate || null;
      needsBootstrap = !!stockTotals.needsBootstrap;
    } else if (typeof getStockTotalsForOutlet_ === 'function') {
      var stockTotalsFallback = getStockTotalsForOutlet_(userOutlet);
      purchaseQty = stockTotalsFallback.purchased || {};
      recordedConsumptionMap = stockTotalsFallback.consumed || {};
      salesConsumptionMap = stockTotalsFallback.sold || {};
      stockSource = stockTotalsFallback.stockSource || stockSource;
      historyAsOfDate = stockTotalsFallback.historyAsOfDate || null;
      needsBootstrap = !!stockTotalsFallback.needsBootstrap;
    } else {
      purchaseQty = _readPurchaseQtyByCode_(userOutlet);
      var bomResult = typeof computeLifetimeBomConsumptionForOutlet_ === 'function'
        ? computeLifetimeBomConsumptionForOutlet_(userOutlet)
        : { consumptionMap: {}, recipeItemCodes: {} };
      salesConsumptionMap = bomResult.consumptionMap || {};
      recordedConsumptionMap = _readRecordedConsumptionQtyByCode_(userOutlet);
      stockSource = 'BOM+Consumption';
    }

    var reorderLookup = _readKeyfigureReorderByCode_(userOutlet);
    var supplierLookup = { byCode: {}, byName: {} };
    buyProducts.forEach(function(p) {
      if (!p) return;
      var sup = (p.supplier || '').toString().trim();
      if (!sup) return;
      if (p.productCode) supplierLookup.byCode[p.productCode] = sup;
      if (p.name) supplierLookup.byName[p.name.toLowerCase()] = sup;
    });
    var recipeItemCodes = {};
    if (!needsBootstrap && typeof buildRecipeItemCodeSet_ === 'function') {
      recipeItemCodes = buildRecipeItemCodeSet_(
        typeof getCachedBomRecipes_ === 'function' ? getCachedBomRecipes_() : (
          typeof loadActiveBomRecipes_ === 'function' ? loadActiveBomRecipes_() : { lines: [] }
        ),
        typeof getCachedBomUsageMatrix_ === 'function' ? getCachedBomUsageMatrix_() : (
          typeof loadActiveBomUsageMatrix_ === 'function' ? loadActiveBomUsageMatrix_() : { lines: [] }
        )
      );
    } else if (!needsBootstrap && typeof computeLifetimeBomConsumptionForOutlet_ === 'function') {
      recipeItemCodes = (computeLifetimeBomConsumptionForOutlet_(userOutlet).recipeItemCodes || {});
    }

    var items = [];
    buyProducts.forEach(function(p) {
      var code = (p.productCode || '').toString().trim();
      if (!code) return;
      var nk = typeof normalizeProductCodeKey_ === 'function' ? normalizeProductCodeKey_(code) : code;
      var totalPurchased = Math.round((purchaseQty[nk] || 0) * 1000) / 1000;
      var totalSold = Math.round((salesConsumptionMap[nk] || 0) * 1000) / 1000;
      var totalConsumed = Math.round((recordedConsumptionMap[nk] || 0) * 1000) / 1000;
      var remaining = Math.round((totalPurchased - totalConsumed) * 1000) / 1000;
      var hasRecipe = !!recipeItemCodes[nk];
      items.push({
        name: p.name || code,
        productCode: code,
        category: p.category || '',
        navCategory: p.navCategory || '',
        unit: p.unit || '',
        totalPurchased: totalPurchased,
        totalConsumed: totalConsumed,
        totalSold: totalSold,
        remaining: remaining,
        reorderQty: reorderLookup[nk] || 0,
        supplier: supplierLookup.byCode[code] || supplierLookup.byName[(p.name || '').toLowerCase()] || '',
        hasRecipe: hasRecipe
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
  } catch (e) {
    Logger.log('Error in getAyamStockQuick: ' + e.toString());
    return { items: [], stockSource: 'BOM', error: e.toString() };
  }
}

/**
 * Inventory Stock Report
 * Calculates remaining stock for sell-type products using:
 *   Stock = (Purchased parent qty × UnitsPerParent) - Sold qty
 * Products sheet must have: ProductCode, ParentProductCode, UnitsPerParent
 */
function getInventoryStockReport(deviceIdOrParams, username) {
  // Support both direct args (deviceId, username) and params object from SecurityWrapper
  var deviceId, resolvedUsername;
  if (deviceIdOrParams && typeof deviceIdOrParams === 'object') {
    deviceId = deviceIdOrParams.deviceId || (deviceIdOrParams._authContext && deviceIdOrParams._authContext.deviceId) || (deviceIdOrParams._validatedUser && deviceIdOrParams._validatedUser.deviceId) || null;
    resolvedUsername = deviceIdOrParams.username || (deviceIdOrParams._authContext && deviceIdOrParams._authContext.username) || (deviceIdOrParams._validatedUser && deviceIdOrParams._validatedUser.username) || null;
  } else {
    deviceId = deviceIdOrParams;
    resolvedUsername = username;
  }
  try {
    const ss = getGlobalSpreadsheet();
    const productsSheet = ss.getSheetByName(SHEETS.PRODUCTS);
    const ordersSheet = ss.getSheetByName(SHEETS.ORDERS);
    const purchasesSheet = ss.getSheetByName(SHEETS.PURCHASES);
    
    if (!productsSheet || !ordersSheet || !purchasesSheet) {
      return { error: 'Required sheets not found' };
    }
    
    // Get user outlet
    var session = getUserSessionEnhanced(deviceId, resolvedUsername) || {};
    var userOutlet = session.outlet;
    if (typeof resolveUserOutlet_ === 'function' && session.username) {
      userOutlet = resolveUserOutlet_(session.username, userOutlet);
    }
    
    // Auto-refresh Stock_History through yesterday on navigation to Stock Report.
    if (typeof ensureStockHistoryCurrentForOutlet_ === 'function') {
      ensureStockHistoryCurrentForOutlet_(userOutlet);
    }
    
    // --- 1. Read Products sheet: build sell items map ---
    const prodData = productsSheet.getDataRange().getValues();
    if (prodData.length < 2) return { items: [], parentSummary: [] };
    const prodHeaders = prodData.shift();
    
    const pNameIdx = prodHeaders.findIndex(h => h && /name/i.test(h.toString()));
    const pCatIdx = prodHeaders.findIndex(h => h && /category/i.test(h.toString()));
    const pTypeIdx = prodHeaders.findIndex(h => h && /type/i.test(h.toString()));
    const pCodeIdx = prodHeaders.findIndex(h => h && /product.*code|^code$/i.test(h.toString()));
    const pParentIdx = prodHeaders.findIndex(h => h && /parent.*product.*code|parent.*code/i.test(h.toString()));
    const pUnitsIdx = prodHeaders.findIndex(h => h && /units.*per.*parent/i.test(h.toString()));
    const pOutletIdx = prodHeaders.findIndex(h => h && /outlet/i.test(h.toString()));
    const pActiveIdx = prodHeaders.findIndex(h => h && /^active$/i.test(h.toString()));
    
    if (pParentIdx === -1 || pUnitsIdx === -1) {
      return { error: 'Products sheet missing ParentProductCode or UnitsPerParent columns' };
    }
    
    // Build sell items map: productCode → { name, category, parentCode, unitsPerParent }
    const sellItems = {};
    prodData.forEach(function(row) {
      var type = pTypeIdx >= 0 ? (row[pTypeIdx] || '').toString().toLowerCase() : '';
      if (type !== 'sell') return;
      if (pActiveIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(row[pActiveIdx])) return;
      
      // Outlet filter
      if (pOutletIdx >= 0 && userOutlet) {
        var outlets = (row[pOutletIdx] || '').toString();
        if (outlets && outlets.toUpperCase() !== 'ALL' && outlets.indexOf(userOutlet) === -1) return;
      }
      
      var code = pCodeIdx >= 0 ? (row[pCodeIdx] || '').toString().trim() : '';
      var parentCode = (row[pParentIdx] || '').toString().trim();
      var unitsPerParent = Number(row[pUnitsIdx]) || 0;
      
      if (!code || !parentCode || unitsPerParent <= 0) return;
      
      sellItems[code] = {
        name: row[pNameIdx] || '',
        category: pCatIdx >= 0 ? (row[pCatIdx] || '') : '',
        parentCode: parentCode,
        unitsPerParent: unitsPerParent
      };
    });
    
    // --- 2. Read Purchases sheet: sum qty by product code ---
    var purchaseQty = {};
    var purchaseData = purchasesSheet.getDataRange().getValues();
    if (purchaseData.length > 1) {
      var phHeaders = purchaseData.shift();
      var phCodeIdx = phHeaders.findIndex(function(h) { return h && /code/i.test(h.toString()); });
      var phQtyIdx = phHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
      var phOutletIdx = phHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
      
      purchaseData.forEach(function(row) {
        if (phOutletIdx !== -1 && userOutlet) {
          if (typeof outletRowMatchesUser_ === 'function') {
            if (!outletRowMatchesUser_(row[phOutletIdx], userOutlet)) return;
          } else if (row[phOutletIdx] !== userOutlet) return;
        }
        var code = phCodeIdx >= 0 ? (row[phCodeIdx] || '').toString().trim() : '';
        var qty = Number(row[phQtyIdx]) || 0;
        if (!code || qty === 0) return;
        purchaseQty[code] = (purchaseQty[code] || 0) + qty;
      });
    }
    
    // --- 3. Read Orders sheet: sum qty sold by product code ---
    var soldQty = {};
    var orderData = ordersSheet.getDataRange().getValues();
    if (orderData.length > 1) {
      var ohHeaders = orderData.shift();
      var ohCodeIdx = ohHeaders.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
      var ohQtyIdx = ohHeaders.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
      var ohOutletIdx = ohHeaders.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
      
      orderData.forEach(function(row) {
        if (ohOutletIdx !== -1 && userOutlet) {
          if (typeof outletRowMatchesUser_ === 'function') {
            if (!outletRowMatchesUser_(row[ohOutletIdx], userOutlet)) return;
          } else if (row[ohOutletIdx] !== userOutlet) return;
        }
        var code = ohCodeIdx >= 0 ? (row[ohCodeIdx] || '').toString().trim() : '';
        var qty = Number(row[ohQtyIdx]) || 0;
        if (!code || qty === 0) return;
        soldQty[code] = (soldQty[code] || 0) + qty;
      });
    }
    
    // --- 4. Calculate stock for each sell item ---
    var items = [];
    var parentTotals = {}; // track parent usage
    
    Object.keys(sellItems).forEach(function(code) {
      var item = sellItems[code];
      var parentPurchased = purchaseQty[item.parentCode] || 0;
      var available = parentPurchased * item.unitsPerParent;
      var sold = soldQty[code] || 0;
      var remaining = available - sold;
      
      // Track parent consumption
      if (!parentTotals[item.parentCode]) {
        parentTotals[item.parentCode] = { purchased: parentPurchased, children: [] };
      }
      parentTotals[item.parentCode].children.push({ code: code, name: item.name, sold: sold, unitsPerParent: item.unitsPerParent });
      
      items.push({
        product: item.name,
        productCode: code,
        category: item.category,
        parentCode: item.parentCode,
        unitsPerParent: item.unitsPerParent,
        parentPurchased: parentPurchased,
        available: available,
        sold: sold,
        remaining: remaining
      });
    });
    
    // Sort by remaining stock ascending (lowest stock first)
    items.sort(function(a, b) { return a.remaining - b.remaining; });

    var ingredients = [];
    if (typeof loadActiveBomRecipes_ === 'function') {
      var bomData = loadActiveBomRecipes_();
      var recipeCodes = {};
      bomData.lines.forEach(function(line) {
        if (line.recipeItemCode) recipeCodes[line.recipeItemCode] = true;
      });

      var consumptionQty = {};
      var consumptionSheet = ss.getSheetByName(SHEETS.INVENTORY_CONSUMPTION);
      if (consumptionSheet && consumptionSheet.getLastRow() > 1) {
        var consData = consumptionSheet.getDataRange().getValues();
        var ch = consData[0];
        var cCodeIdx = ch.findIndex(function(h) { return h && /product.*code|^code$/i.test(h.toString()); });
        var cQtyIdx = ch.findIndex(function(h) { return h && /qty|quantity/i.test(h.toString()); });
        var cOutletIdx = ch.findIndex(function(h) { return h && /outlet/i.test(h.toString()); });
        for (var ci = 1; ci < consData.length; ci++) {
          var crow = consData[ci];
          if (cOutletIdx !== -1 && userOutlet) {
            if (typeof outletRowMatchesUser_ === 'function') {
              if (!outletRowMatchesUser_(crow[cOutletIdx], userOutlet)) continue;
            } else if (crow[cOutletIdx] !== userOutlet) continue;
          }
          var cCode = cCodeIdx >= 0 ? (crow[cCodeIdx] || '').toString().trim() : '';
          var cQty = cQtyIdx >= 0 ? (Number(crow[cQtyIdx]) || 0) : 0;
          if (!cCode || cQty === 0) continue;
          consumptionQty[cCode] = (consumptionQty[cCode] || 0) + cQty;
        }
      }

      var productNames = typeof getProductNamesByCode === 'function' ? getProductNamesByCode() : {};
      Object.keys(recipeCodes).forEach(function(code) {
        var purchased = purchaseQty[code] || 0;
        var consumed = consumptionQty[code] || 0;
        ingredients.push({
          productCode: code,
          product: productNames[code] || code,
          purchased: purchased,
          consumed: consumed,
          remaining: purchased - consumed
        });
      });
      ingredients.sort(function(a, b) { return a.remaining - b.remaining; });
    }
    
    return { items: items, ingredients: ingredients };
    
  } catch (error) {
    Logger.log("Error in getInventoryStockReport: " + error.toString());
    return { error: error.toString() };
  }
}
