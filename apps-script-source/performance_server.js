/**
 * Performance-optimized server functions
 * This file contains optimized versions of commonly called functions
 */

// Cached spreadsheet connections
var _cachedSpreadsheet = null;
var _cachedSheets = {};

/**
 * Get cached spreadsheet connection
 */
function getCachedSpreadsheet() {
  if (!_cachedSpreadsheet) {
    _cachedSpreadsheet = getGlobalSpreadsheet();
  }
  return _cachedSpreadsheet;
}

/**
 * Get cached sheet by name
 */
function getCachedSheet(sheetName) {
  if (!_cachedSheets[sheetName]) {
    _cachedSheets[sheetName] = getCachedSpreadsheet().getSheetByName(sheetName);
  }
  return _cachedSheets[sheetName];
}

/**
 * Batch data loader - loads multiple types of data in a single call
 */
function batchLoadFormData() {
  try {
    Logger.log("=== Batch Loading Form Data ===");
    
    const result = {
      products: null,
      expenseAccounts: null,
      investmentAccounts: null,
      paymentSources: null,
      outlets: null
    };
    
    // Load products
    try {
      result.products = getProducts('Sell');
    } catch (error) {
      Logger.log("Error loading products: " + error.toString());
      result.products = "[]";
    }
    
    // Load expense accounts
    try {
      result.expenseAccounts = getExpenseAccounts();
    } catch (error) {
      Logger.log("Error loading expense accounts: " + error.toString());
      result.expenseAccounts = "[]";
    }
    
    // Load investment accounts
    try {
      result.investmentAccounts = getInvestmentAccounts();
    } catch (error) {
      Logger.log("Error loading investment accounts: " + error.toString());
      result.investmentAccounts = "[]";
    }
    
    // Load payment sources
    try {
      result.paymentSources = getPaymentSources();
    } catch (error) {
      Logger.log("Error loading payment sources: " + error.toString());
      result.paymentSources = "[]";
    }
    
    // Load outlets
    try {
      result.outlets = getOutlets();
    } catch (error) {
      Logger.log("Error loading outlets: " + error.toString());
      result.outlets = "[]";
    }
    
    Logger.log("Batch load completed successfully");
    return JSON.stringify(result);
    
  } catch (error) {
    Logger.log("Error in batchLoadFormData: " + error.toString());
    return JSON.stringify({
      products: "[]",
      expenseAccounts: "[]", 
      investmentAccounts: "[]",
      paymentSources: "[]",
      outlets: "[]"
    });
  }
}

/**
 * Optimized products function with better caching
 */
function getProductsOptimized(type) {
  try {
    Logger.log("=== Getting Products (Optimized) ===");
    
    // Use cached sheet
    const sheet = getCachedSheet(SHEETS.PRODUCTS);
    if (!sheet) {
      Logger.log("Products sheet not found");
      return "[]";
    }
    
    // Get all data at once instead of range by range
    const data = sheet.getDataRange().getValues();
    Logger.log("Retrieved " + data.length + " rows from Products sheet");
    
    if (data.length < 2) {
      Logger.log("No data rows found");
      return "[]";
    }
    
    const headers = data[0];
    const nameIdx = headers.findIndex(h => /name/i.test(h));
    const priceIdx = headers.findIndex(h => /price/i.test(h));
    const categoryIdx = headers.findIndex(h => /category/i.test(h));
    const typeIdx = headers.findIndex(h => /type/i.test(h));
    const activeIdx = headers.findIndex(h => /^active$/i.test(h));
    
    if (nameIdx === -1 || priceIdx === -1 || categoryIdx === -1) {
      Logger.log("Required columns not found");
      return "[]";
    }
    
    // Process all rows in memory instead of individual sheet operations
    const products = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(row[activeIdx])) continue;
      const name = row[nameIdx];
      const price = row[priceIdx];
      const category = row[categoryIdx];
      const productType = typeIdx !== -1 ? row[typeIdx] : 'Sell';
      
      if (name && price && category) {
        // Filter by type if specified
        if (!type || productType === type) {
          products.push({
            Name: name,
            Price: price,
            Category: category,
            Type: productType || 'Sell'
          });
        }
      }
    }
    
    Logger.log("Returning " + products.length + " products");
    return JSON.stringify(products);
    
  } catch (error) {
    Logger.log("Error in getProductsOptimized: " + error.toString());
    return "[]";
  }
}

/**
 * Optimized accounts loading
 */
function getAccountsOptimized(accountType) {
  try {
    const sheet = getCachedSheet(SHEETS.ACCOUNTS);
    if (!sheet) {
      Logger.log("Accounts sheet not found");
      return "[]";
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return "[]";
    }
    
    const headers = data[0];
    const nameIdx = headers.findIndex(h => /name/i.test(h));
    const typeIdx = headers.findIndex(h => /type/i.test(h));
    
    const accounts = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[nameIdx];
      const type = typeIdx !== -1 ? row[typeIdx] : '';
      
      if (name && (!accountType || type === accountType)) {
        accounts.push(name);
      }
    }
    
    return JSON.stringify(accounts);
    
  } catch (error) {
    Logger.log("Error in getAccountsOptimized: " + error.toString());
    return "[]";
  }
}

/**
 * Clear cached connections (call this when structure changes)
 */
function clearServerCache() {
  _cachedSpreadsheet = null;
  _cachedSheets = {};
  if (typeof clearTabAccessCache_ === 'function') {
    clearTabAccessCache_();
  }
  Logger.log("Server cache cleared");
}

/** @deprecated Use clearServerCache — kept for serverSecurity.js call sites */
function clearSheetCache() {
  clearServerCache();
}

/**
 * Optimized sales summary with better date handling
 */
function getSalesSummaryOptimized(dateString, userEmail) {
  try {
    Logger.log("=== Getting Sales Summary (Optimized) ===");
    Logger.log("Date: " + dateString + ", User: " + userEmail);
    
    // Get user outlet
    let userOutlet = null;
    if (userEmail) {
      const session = getUserSessionEnhanced(null, userEmail) || {};
      userOutlet = session.outlet;
    }
    
    const sheet = getCachedSheet(SHEETS.ORDERS);
    if (!sheet) {
      return JSON.stringify({ success: false, message: "Orders sheet not found" });
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return JSON.stringify({ success: true, data: [], summary: emptySalesSummary(dateString) });
    }
    
    const headers = data[0];
    const dateIdx = headers.findIndex(h => /timestamp|date/i.test(h));
    const productIdx = headers.findIndex(h => /product|name/i.test(h));
    const qtyIdx = headers.findIndex(h => /qty|quantity/i.test(h));
    const totalIdx = headers.findIndex(h => /total|amount/i.test(h));
    const outletIdx = headers.findIndex(h => /outlet/i.test(h));
    
    // Parse target date once
    const targetDate = new Date(dateString);
    const targetDateStr = targetDate.getFullYear() + '-' + 
                         String(targetDate.getMonth() + 1).padStart(2, '0') + '-' + 
                         String(targetDate.getDate()).padStart(2, '0');
    
    // Process all data in memory
    const salesData = [];
    const productSummary = {};
    let totalRevenue = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Filter by outlet
      if (userOutlet && outletIdx >= 0 && row[outletIdx] !== userOutlet) continue;
      
      // Check date
      const orderDate = row[dateIdx];
      if (!orderDate) continue;
      
      let orderDateStr;
      if (orderDate instanceof Date) {
        orderDateStr = orderDate.getFullYear() + '-' + 
                      String(orderDate.getMonth() + 1).padStart(2, '0') + '-' + 
                      String(orderDate.getDate()).padStart(2, '0');
      } else {
        orderDateStr = String(orderDate).split(' ')[0]; // Take just the date part
      }
      
      if (orderDateStr === targetDateStr) {
        const product = row[productIdx];
        const qty = Number(row[qtyIdx]) || 0;
        const total = Number(row[totalIdx]) || 0;
        
        if (product && qty > 0) {
          salesData.push({
            product: product,
            quantity: qty,
            total: total
          });
          
          // Update summary
          if (!productSummary[product]) {
            productSummary[product] = { quantity: 0, total: 0 };
          }
          productSummary[product].quantity += qty;
          productSummary[product].total += total;
          totalRevenue += total;
        }
      }
    }
    
    const summary = {
      date: dateString,
      totalRevenue: totalRevenue,
      totalItems: Object.keys(productSummary).length,
      productSummary: productSummary
    };
    
    Logger.log("Sales summary completed - Revenue: " + totalRevenue + ", Items: " + salesData.length);
    return JSON.stringify({ success: true, data: salesData, summary: summary });
    
  } catch (error) {
    Logger.log("Error in getSalesSummaryOptimized: " + error.toString());
    return JSON.stringify({ success: false, message: error.toString() });
  }
}
