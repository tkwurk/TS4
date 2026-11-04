// Server-side authentication enforcement
// Add these functions to your main server files (code.js, orders.js, etc.)

/**
 * Validates authentication context from client calls
 * Call this at the beginning of every server function that requires authentication
 */
function validateAuthenticationContext(authContext) {
  try {
    if (!authContext || !authContext.deviceId || !authContext.username) {
      throw new Error('Authentication required - incomplete context');
    }
    if (authContext.timestamp) {
      var requestAge = new Date().getTime() - authContext.timestamp;
      if (requestAge > 5 * 60 * 1000) {
        throw new Error('Authentication expired - request too old');
      }
    }
    var session = requireAuth_({ _authContext: authContext });
    var outlet = session.outlet;
    if (typeof resolveUserOutlet_ === 'function') {
      outlet = resolveUserOutlet_(session.username, outlet);
    }
    return {
      valid: true,
      username: session.username,
      outlet: outlet,
      allowedTabs: session.allowedTabs || resolveEffectiveTabsFromSession(session),
      role: session.role || 'cashier',
      deviceId: authContext.deviceId
    };
  } catch (error) {
    Logger.log("SECURITY: Authentication validation failed: " + error.toString());
    throw error;
  }
}

function requireAuth(params) {
  return requireAuth_(params);
}

function requirePermission(params, permission) {
  return requirePermission_(params, permission);
}

function requireAdmin(params) {
  return requireAdmin_(params);
}

/**
 * Secure wrapper for server functions that require authentication
 * Use this to wrap existing server functions
 */
function requireAuthentication(originalFunction) {
  return function(params) {
    try {
      // Extract authentication context
      const authContext = params && params._authContext;
      
      // Validate authentication
      const authResult = validateAuthenticationContext(authContext);
      
      // Remove auth context from params before passing to original function
      const cleanParams = { ...params };
      delete cleanParams._authContext;
      
      // Add validated user info to params
      cleanParams._validatedUser = authResult;
      
      // Call original function with validated context
      return originalFunction.call(this, cleanParams);
      
    } catch (error) {
      Logger.log("SECURITY: Authentication failed for server function: " + error.toString());
      throw error;
    }
  };
}

/**
 * Secure version of getProducts function
 */
function getProductsSecure(params) {
  const authResult = validateAuthenticationContext(params._authContext);
  const typeFilter = params.type || 'Sell';
  const getAllTypes = typeFilter === 'ALL_TYPES';

  Logger.log("SECURITY: Authenticated getProducts call for user: " + authResult.username + ", outlet: " + authResult.outlet);

  try {
    var outlet = authResult.outlet;
    if (typeof resolveUserOutlet_ === 'function') {
      outlet = resolveUserOutlet_(authResult.username, outlet);
    }
    if (!outlet || outlet === 'undefined') {
      Logger.log("SECURITY: No outlet on session for " + authResult.username + " — returning universal/outlet-less products only");
    }
    if (typeof getProducts !== 'function') {
      throw new Error('getProducts not available');
    }
    var productsJson = getProducts(getAllTypes ? null : typeFilter, outlet || '');
    Logger.log("SECURITY: getProducts returned for outlet: " + (outlet || '(none)'));
    return productsJson;
  } catch (error) {
    Logger.log("SECURITY: Error in getProductsSecure: " + error.toString());
    throw error;
  }
}

/**
 * Secure version of getExpenseAccounts function
 */
function getExpenseAccountsSecure(params) {
  const authResult = validateAuthenticationContext(params._authContext);
  
  Logger.log("SECURITY: Authenticated getExpenseAccounts call for user: " + authResult.username);
  
  try {
    const sheet = getCachedSheet(SHEETS.ACCOUNTS);
    if (!sheet) {
      throw new Error("Accounts sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const accounts = data.slice(1)
      .filter(row => row[1] && row[1].toString().toLowerCase() === 'expense')
      .map(row => row[0])
      .filter(Boolean);
    
    Logger.log("SECURITY: Returning " + accounts.length + " expense accounts");
    return accounts;
    
  } catch (error) {
    Logger.log("SECURITY: Error in getExpenseAccountsSecure: " + error.toString());
    throw error;
  }
}

/**
 * Secure version that returns ALL account names (any type). Used by purchasing so
 * the account dropdown lists every account; the backend derives whether a purchase
 * is an investment from the selected account's type.
 */
function getAllAccountsSecure(params) {
  const authResult = validateAuthenticationContext(params._authContext);

  Logger.log("SECURITY: Authenticated getAllAccounts call for user: " + authResult.username);

  try {
    const sheet = getCachedSheet(SHEETS.ACCOUNTS);
    if (!sheet) {
      throw new Error("Accounts sheet not found");
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const accounts = data.slice(1)
      .map(row => row[0])
      .filter(Boolean);
    Logger.log("SECURITY: Returning " + accounts.length + " accounts");
    return accounts;
  } catch (error) {
    Logger.log("SECURITY: Error in getAllAccountsSecure: " + error.toString());
    throw error;
  }
}

/**
 * Secure version of getPaymentSources function
 */
function getPaymentSourcesSecure(params) {
  const authResult = validateAuthenticationContext(params._authContext);
  
  Logger.log("SECURITY: Authenticated getPaymentSources call for user: " + authResult.username);
  
  try {
    const sheet = getCachedSheet(SHEETS.PAYMENT_SOURCES);
    if (!sheet) {
      throw new Error("Payment Sources sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const sources = data.slice(1)
      .map(row => row[0])
      .filter(Boolean);
    
    Logger.log("SECURITY: Returning " + sources.length + " payment sources");
    return sources;
    
  } catch (error) {
    Logger.log("SECURITY: Error in getPaymentSourcesSecure: " + error.toString());
    throw error;
  }
}

/**
 * Secure version of getDepositAccounts function
 */
function getDepositAccountsSecure(params) {
  const authResult = validateAuthenticationContext(params._authContext);
  
  Logger.log("SECURITY: Authenticated getDepositAccounts call for user: " + authResult.username);
  
  try {
    const sheet = getCachedSheet(SHEETS.ACCOUNTS);
    if (!sheet) {
      throw new Error("Accounts sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const accounts = data.slice(1)
      .filter(row => row[1] && row[1].toString().toLowerCase() === 'deposit')
      .map(row => row[0])
      .filter(Boolean);
    
    Logger.log("SECURITY: Returning " + accounts.length + " deposit accounts");
    return accounts;
    
  } catch (error) {
    Logger.log("SECURITY: Error in getDepositAccountsSecure: " + error.toString());
    throw error;
  }
}

/**
 * Clears the script-level cache for all users.
 * This function should be called by an administrator when a global cache reset is needed.
 * For example, after updating the Products or Outlets sheets.
 *
 * @param {object} params - The parameters object from the client.
 * @returns {object} A result object indicating success or failure.
 */
function clearAllUsersCacheSecure(params) {
  // First, validate that the user has administrative privileges
  const authResult = validateAuthenticationContext(params._authContext);
  
  // Check if the user is an admin (assuming 'admin' is a special tab or role)
  const isAdmin = authResult.allowedTabs && (
    authResult.allowedTabs.includes('admin') ||
    authResult.allowedTabs.includes('adminOrders') ||
    authResult.allowedTabs.includes('*')
  );
  
  if (!isAdmin) {
    Logger.log("SECURITY: Non-admin user '" + authResult.username + "' attempted to clear global cache.");
    return { success: false, error: "You do not have permission to perform this action." };
  }
  
  Logger.log("ADMIN ACTION: User '" + authResult.username + "' is clearing the global script cache.");
  
  try {
    // Clear sheet cache
    if (typeof clearSheetCache === 'function') {
      clearSheetCache();
      Logger.log("Sheet cache cleared successfully.");
    }
    
    // If you were using CacheService directly, you would do this:
    // const scriptCache = CacheService.getScriptCache();
    // scriptCache.removeAll(['some_key_to_remove']);

    return { success: true, message: "Global server cache has been cleared for all users." };
    
  } catch (error) {
    Logger.log("Error clearing global cache: " + error.toString());
    return { success: false, error: "An error occurred while clearing the cache." };
  }
}

/**
 * Secure wrapper to fetch all products (all types) and unique category/navcategory lists for admin use.
 * Requires admin tab permission.
 */
function getProductsForAdminSecure(params) {
  var authResult = validateAuthenticationContext(params._authContext);
  var isAdmin = authResult.allowedTabs && (authResult.allowedTabs.includes('adminOrders') || authResult.allowedTabs.includes('*'));
  if (!isAdmin) {
    return { success: false, error: 'Admin permission required.' };
  }

  try {
    if (typeof ensureProductActiveColumn === 'function') ensureProductActiveColumn();
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return { success: false, error: 'Products sheet not found.' };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, products: [], categories: [], navCategories: [], suppliers: [] };

    var headers = data[0].map(function(h) { return h.toString().trim(); });
    var catIdx  = headers.findIndex(function(h) { return h.toUpperCase() === 'CATEGORY'; });
    var navIdx  = headers.findIndex(function(h) { return h.toUpperCase() === 'NAVCATEGORY'; });
    var supIdx  = headers.findIndex(function(h) { return h.toUpperCase() === 'SUPPLIER'; });

    var products = [];
    var catSet = {}, navSet = {}, supSet = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });
      products.push(obj);

      var cat = catIdx >= 0 ? (row[catIdx] || '').toString().trim() : '';
      var nav = navIdx >= 0 ? (row[navIdx] || '').toString().trim() : '';
      var sup = supIdx >= 0 ? (row[supIdx] || '').toString().trim() : '';
      if (cat) catSet[cat] = true;
      if (nav) navSet[nav] = true;
      if (sup) supSet[sup] = true;
    }

    return {
      success: true,
      products: products,
      categories: Object.keys(catSet).sort(),
      navCategories: Object.keys(navSet).sort(),
      suppliers: Object.keys(supSet).sort()
    };
  } catch (e) {
    Logger.log('getProductsForAdminSecure error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Secure function to save (create) a new product row in the Products sheet.
 * Validates uniqueness of PRODUCTCODE before inserting.
 * Requires admin tab permission.
 */
function saveProductSecure(params) {
  try {
    var session = requireAdmin_(params);
    var product = params.product;
    if (!product) return apiFail('No product data provided.', 'PRODUCT_INVALID');

    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return apiFail('Products sheet not found.', 'PRODUCTS_SHEET_MISSING');

    if (typeof ensureProductCookingColumns === 'function') ensureProductCookingColumns();
    if (typeof ensureProductActiveColumn === 'function') ensureProductActiveColumn();

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim(); });

    var codeIdx = headers.findIndex(function(h) { return h.toUpperCase() === 'PRODUCTCODE'; });
    var typeIdx = headers.findIndex(function(h) { return h.toUpperCase() === 'TYPE'; });
    if (codeIdx >= 0 && product.PRODUCTCODE) {
      var code = product.PRODUCTCODE.toString().trim().toUpperCase();
      var newType = (product.TYPE || '').toString().trim().toUpperCase();
      for (var i = 1; i < data.length; i++) {
        var existing = (data[i][codeIdx] || '').toString().trim().toUpperCase();
        if (existing === code) {
          var existingType = typeIdx >= 0 ? (data[i][typeIdx] || '').toString().trim().toUpperCase() : '';
          if (!newType || !existingType || newType === existingType) {
            return apiFail('A ' + (newType || 'product') + ' with code "' + code + '" already exists.', 'PRODUCT_DUPLICATE');
          }
        }
      }
    }

    var newRow = headers.map(function(h) {
      if (product.hasOwnProperty(h)) return product[h];
      var key = Object.keys(product).find(function(k) { return k.trim().toUpperCase() === h.trim().toUpperCase(); });
      return key !== undefined ? product[key] : '';
    });

    var activeIdx = headers.findIndex(function(h) { return h.trim().toUpperCase() === 'ACTIVE'; });
    if (activeIdx >= 0 && !(newRow[activeIdx] || '').toString().trim()) {
      newRow[activeIdx] = 'Y';
    }

    sheet.appendRow(newRow);
    if (typeof writeAuditLog_ === 'function') {
      writeAuditLog_(session.username, 'CREATE_PRODUCT', product.PRODUCTCODE || product.NAME, JSON.stringify(product), session.outlet);
    }
    if (typeof clearSheetCache === 'function') clearSheetCache();

    return apiOk({ productCode: product.PRODUCTCODE }, 'PRODUCT_CREATED');
  } catch (e) {
    Logger.log('saveProductSecure error: ' + e.toString());
    return apiFail(e.message || e.toString(), 'PRODUCT_SAVE_FAILED');
  }
}

/**
 * Secure function to update Name and Price fields of an existing product row.
 * Finds the row by PRODUCTCODE + TYPE, updates only those two fields.
 * Requires admin tab permission.
 */
function updateProductFieldsSecure(params) {
  try {
    var session = requireAdmin_(params);
    var productCode  = (params.productCode || '').toString().trim().toUpperCase();
    var productType  = (params.productType || '').toString().trim().toUpperCase();
    var newName      = params.name !== undefined && params.name !== null ? params.name.toString().trim() : null;
    var newPrice     = (params.price !== null && params.price !== undefined && !isNaN(parseFloat(params.price))) ? parseFloat(params.price) : null;
    var softDelete   = params.active === 'N' || params.softDelete === true;
    var setActive    = params.active !== undefined && params.active !== null;
    var activeVal    = setActive ? params.active.toString().trim().toUpperCase() : null;

    if (!productCode) return apiFail('productCode is required.', 'PRODUCT_INVALID');

    if (typeof ensureProductActiveColumn === 'function') ensureProductActiveColumn();
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return apiFail('Products sheet not found.', 'PRODUCTS_SHEET_MISSING');

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim().toUpperCase(); });

    var codeIdx      = headers.indexOf('PRODUCTCODE');
    var typeIdx      = headers.indexOf('TYPE');
    var nameIdx      = headers.indexOf('NAME');
    var priceIdx     = headers.indexOf('PRICE');
    var buyPriceIdx  = headers.indexOf('BUYPRICE');
    var imageIdx     = headers.indexOf('IMAGEFILEID');
    var activeIdx    = headers.indexOf('ACTIVE');

    var newImage = (params.imageFileId !== undefined && params.imageFileId !== null) ? params.imageFileId.toString().trim() : null;

    if (codeIdx < 0) return apiFail('PRODUCTCODE column not found.', 'PRODUCT_SCHEMA');

    for (var i = 1; i < data.length; i++) {
      var rowCode = (data[i][codeIdx] || '').toString().trim().toUpperCase();
      var rowType = typeIdx >= 0 ? (data[i][typeIdx] || '').toString().trim().toUpperCase() : '';
      if (rowCode !== productCode) continue;
      if (productType && rowType && rowType !== productType) continue;

      var rowNum = i + 1;
      if (newName !== null) sheet.getRange(rowNum, nameIdx + 1).setValue(newName);
      if (newPrice !== null) {
        var isBuy = rowType === 'BUY';
        if (isBuy && buyPriceIdx >= 0) sheet.getRange(rowNum, buyPriceIdx + 1).setValue(newPrice);
        else if (!isBuy && priceIdx >= 0) sheet.getRange(rowNum, priceIdx + 1).setValue(newPrice);
      }
      if (newImage !== null && imageIdx >= 0) sheet.getRange(rowNum, imageIdx + 1).setValue(newImage);
      if (activeIdx >= 0 && (softDelete || setActive)) {
        if (softDelete || activeVal === 'N') {
          sheet.getRange(rowNum, activeIdx + 1).setValue('N');
        } else if (activeVal === 'Y') {
          sheet.getRange(rowNum, activeIdx + 1).setValue('Y');
        }
      }

      var auditAction = softDelete ? 'SOFT_DELETE_PRODUCT' : (setActive && activeVal === 'Y' ? 'ACTIVATE_PRODUCT' : 'UPDATE_PRODUCT');
      if (typeof writeAuditLog_ === 'function') {
        writeAuditLog_(session.username, auditAction, productCode, JSON.stringify(params), session.outlet);
      }
      if (typeof clearSheetCache === 'function') clearSheetCache();
      var msgKey = softDelete ? 'PRODUCT_DEACTIVATED' : (setActive && activeVal === 'Y' ? 'PRODUCT_ACTIVATED' : 'PRODUCT_UPDATED');
      return apiOk({ productCode: productCode }, msgKey);
    }
    return apiFail('Product not found: ' + productCode, 'PRODUCT_NOT_FOUND');
  } catch (e) {
    Logger.log('updateProductFieldsSecure error: ' + e.toString());
    return apiFail(e.message || e.toString(), 'PRODUCT_UPDATE_FAILED');
  }
}

function deactivateProductSecure(params) {
  params = params || {};
  params.softDelete = true;
  params.active = 'N';
  return updateProductFieldsSecure(params);
}

function activateProductSecure(params) {
  params = params || {};
  params.active = 'Y';
  delete params.softDelete;
  return updateProductFieldsSecure(params);
}

/**
 * Initialize secure versions of all server functions
 * Call this at the start of your main server file
 */
function initializeSecureServerFunctions() {
  // Replace original functions with secure versions
  if (typeof getProducts !== 'undefined') {
    window.getProductsOriginal = getProducts;
    getProducts = getProductsSecure;
  }
  
  if (typeof getExpenseAccounts !== 'undefined') {
    window.getExpenseAccountsOriginal = getExpenseAccounts;
    getExpenseAccounts = getExpenseAccountsSecure;
  }
  
  if (typeof getPaymentSources !== 'undefined') {
    window.getPaymentSourcesOriginal = getPaymentSources;
    getPaymentSources = getPaymentSourcesSecure;
  }
  
  if (typeof getDepositAccounts !== 'undefined') {
    window.getDepositAccountsOriginal = getDepositAccounts;
    getDepositAccounts = getDepositAccountsSecure;
  }
  
  if (typeof recordExpense !== 'undefined') {
    window.recordExpenseOriginal = recordExpense;
    recordExpense = recordExpenseSecure;
  }
  
  Logger.log("SECURITY: Secure server functions initialized");
}

/**
 * Middleware to automatically add authentication to server functions
 * Use this for existing functions that need security
 */
function addAuthenticationToExistingFunctions() {
  const functionsToSecure = [
    'getProducts',
    'getExpenseAccounts', 
    'getPaymentSources',
    'getDepositAccounts',
    'recordExpense',
    'recordDeposit',
    'saveOrder',
    'getMonthlyUsageReport',
    'getExpenseHistory',
    'getDepositHistory',
    'getInventoryStockReport',
    'getAyamStockQuick',
    'getMonthlyCashSummary',
    'getDailyCashSummary'
  ];
  
  functionsToSecure.forEach(funcName => {
    if (typeof globalThis[funcName] === 'function') {
      const originalFunc = globalThis[funcName];
      globalThis[funcName] = function(params) {
        // Add authentication validation
        const authResult = validateAuthenticationContext(params._authContext);
        
        // Call original function with validated context
        const cleanParams = { ...params };
        delete cleanParams._authContext;
        cleanParams._validatedUser = authResult;
        
        return originalFunc.call(this, cleanParams);
      };
      
      Logger.log("SECURITY: Added authentication to function: " + funcName);
    }
  });
}

/**
 * Security audit function to check for unauthorized calls
 */
function auditServerSecurity() {
  Logger.log("=== SECURITY AUDIT ===");
  
  const securedFunctions = [
    'getProducts',
    'getExpenseAccounts',
    'getPaymentSources', 
    'getDepositAccounts',
    'recordExpense',
    'recordDeposit',
    'saveOrder'
  ];
  
  securedFunctions.forEach(funcName => {
    if (typeof globalThis[funcName] === 'function') {
      Logger.log("✓ Function secured: " + funcName);
    } else {
      Logger.log("⚠ Function not found: " + funcName);
    }
  });
  
  Logger.log("=== SECURITY AUDIT COMPLETE ===");
  
  return {
    securedFunctions: securedFunctions.filter(fn => typeof globalThis[fn] === 'function'),
    timestamp: new Date().toISOString()
  };
}
