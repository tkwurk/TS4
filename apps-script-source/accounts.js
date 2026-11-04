function getExpenseAccounts() {
  return getAccountsByType('Expense');
}

function getPaymentMethods(deviceId, username) {
  try {
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PAYMENT_METHOD);
    if (!sheet) {
      Logger.log("PaymentMethod sheet not found");
      return [];
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var cols = buildColumnMap(data[0], 'PAYMENT_METHOD');
    return data.slice(1)
      .filter(function(row) {
        var status = row[cols.STATUS];
        return status && status.toString().trim().toLowerCase() === 'active';
      })
      .map(function(row) {
        var nameCell = row[cols.NAME];
        var pnCell = row[cols.PN];
        if (nameCell === undefined || nameCell === null || String(nameCell).trim() === '') return null;
        var name = String(nameCell).trim();
        return {
          pn:                 pnCell != null ? String(pnCell).trim() : '',
          name:               name,
          destinationAccount: row[cols.DESTINATION_ACCOUNT] != null ? String(row[cols.DESTINATION_ACCOUNT]).trim() : '',
          status:             row[cols.STATUS] != null ? String(row[cols.STATUS]).trim() : '',
          outlet:             row[cols.OUTLET] ? row[cols.OUTLET].toString().trim() : '',
          module:             row[cols.MODULE] ? row[cols.MODULE].toString().trim() : '',
          openDrawer:         parsePaymentMethodOpenDrawer(row, cols, name),
          autoDeposit:        parsePaymentMethodAutoDeposit(row, cols, name)
        };
      })
      .filter(function(m) { return m && m.pn && m.name && !isCheckoutOnlyPaymentMethodName_(m.name); });
  } catch (error) {
    Logger.log("Error in getPaymentMethods: " + error.toString());
    return [];
  }
}

/** Active payment methods keyed by lowercase name (server-side reporting + auto-deposit). */
var PAYMENT_METHOD_CONFIG_CACHE_KEY = 'payment_method_config_v1';
var PAYMENT_METHOD_CONFIG_TTL_SEC = 300;

function invalidatePaymentMethodConfigCache_() {
  try {
    CacheService.getScriptCache().remove(PAYMENT_METHOD_CONFIG_CACHE_KEY);
  } catch (e) {
    Logger.log('invalidatePaymentMethodConfigCache_ warning: ' + e.toString());
  }
}

function loadPaymentMethodConfigMap_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(PAYMENT_METHOD_CONFIG_CACHE_KEY);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    Logger.log('loadPaymentMethodConfigMap_ cache read warning: ' + e.toString());
  }

  var map = {};
  try {
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PAYMENT_METHOD);
    if (!sheet) return map;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return map;
    var cols = buildColumnMap(data[0], 'PAYMENT_METHOD');
    data.slice(1).forEach(function(row) {
      var status = row[cols.STATUS];
      if (!status || status.toString().trim().toLowerCase() !== 'active') return;
      var name = row[cols.NAME] ? row[cols.NAME].toString().trim() : '';
      if (!name) return;
      var key = name.toLowerCase();
      map[key] = {
        name: name,
        destinationAccount: row[cols.DESTINATION_ACCOUNT] ? row[cols.DESTINATION_ACCOUNT].toString().trim() : '',
        autoDeposit: parsePaymentMethodAutoDeposit(row, cols, name),
        openDrawer: parsePaymentMethodOpenDrawer(row, cols, name)
      };
    });
    try {
      CacheService.getScriptCache().put(
        PAYMENT_METHOD_CONFIG_CACHE_KEY,
        JSON.stringify(map),
        PAYMENT_METHOD_CONFIG_TTL_SEC
      );
    } catch (cacheWriteErr) {
      Logger.log('loadPaymentMethodConfigMap_ cache write warning: ' + cacheWriteErr.toString());
    }
  } catch (error) {
    Logger.log('loadPaymentMethodConfigMap_ error: ' + error.toString());
  }
  return map;
}

function isCheckoutOnlyPaymentMethodName_(methodName) {
  var key = (methodName || '').toString().trim().toLowerCase();
  return key === 'partial' || key === 'split';
}

function isCashLikePaymentMethodName_(methodName) {
  var key = (methodName || '').toString().trim().toLowerCase();
  return key === 'cash' || key === 'tunai';
}

/** Yes/Y/True/1 in Auto Deposit column; legacy default mirrors old non-cash auto-deposit behaviour. */
function parsePaymentMethodAutoDeposit(row, cols, methodName) {
  var nameLower = String(methodName || '').trim().toLowerCase();
  if (isCheckoutOnlyPaymentMethodName_(nameLower) || isCashLikePaymentMethodName_(nameLower) || nameLower === 'free') {
    return false;
  }
  var idx = cols.AUTO_DEPOSIT;
  if (idx === undefined || idx < 0) {
    return true;
  }
  var raw = row[idx];
  if (raw === undefined || raw === null || String(raw).trim() === '') {
    return true;
  }
  var v = String(raw).trim().toLowerCase();
  if (v === 'no' || v === 'n' || v === 'false' || v === '0') return false;
  return v === 'yes' || v === 'y' || v === 'true' || v === '1';
}

function paymentMethodHasAutoDeposit_(methodName, configMap) {
  var key = (methodName || '').toString().trim().toLowerCase();
  if (!key || isCheckoutOnlyPaymentMethodName_(key) || isCashLikePaymentMethodName_(key) || key === 'free') {
    return false;
  }
  configMap = configMap || loadPaymentMethodConfigMap_();
  if (configMap[key]) return !!configMap[key].autoDeposit;
  return true;
}

function canonicalPaymentMethodName_(methodName, configMap) {
  var key = (methodName || '').toString().trim().toLowerCase();
  configMap = configMap || loadPaymentMethodConfigMap_();
  if (configMap[key] && configMap[key].name) return configMap[key].name;
  if (isCashLikePaymentMethodName_(key)) return 'Cash';
  return (methodName || '').toString().trim() || 'Other';
}

/** Yes/Y/True/1 in Open Drawer column; blank defaults to Cash-only for legacy sheets. */
function parsePaymentMethodOpenDrawer(row, cols, methodName) {
  var idx = cols.OPEN_DRAWER;
  if (idx === undefined || idx < 0) {
    return String(methodName || '').trim().toLowerCase() === 'cash';
  }
  var raw = row[idx];
  if (raw === undefined || raw === null || String(raw).trim() === '') {
    return String(methodName || '').trim().toLowerCase() === 'cash';
  }
  var v = String(raw).trim().toLowerCase();
  return v === 'yes' || v === 'y' || v === 'true' || v === '1';
}

function getDepositAccounts() {
  return getAccountsByType('Asset');
}

var ACCOUNT_TYPE_MAP_CACHE_KEY = 'account_type_map_v1';
var ACCOUNT_TYPE_MAP_TTL_SEC = 300;

function normalizeAccountNameKey_(name) {
  return String(name || '').trim().toLowerCase();
}

function invalidateAccountTypeMapCache_() {
  try {
    CacheService.getScriptCache().remove(ACCOUNT_TYPE_MAP_CACHE_KEY);
  } catch (e) {
    Logger.log('invalidateAccountTypeMapCache_ warning: ' + e.toString());
  }
}

/** @returns {Object<string, string>} lowercase account name → type */
function loadAccountTypeByNameMap_() {
  try {
    var cached = CacheService.getScriptCache().get(ACCOUNT_TYPE_MAP_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) {
    Logger.log('loadAccountTypeByNameMap_ cache read warning: ' + e.toString());
  }

  var map = {};
  try {
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.ACCOUNTS);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var cols = buildColumnMap(data[0], 'ACCOUNTS');
        var nameIdx = cols.NAME;
        var typeIdx = cols.TYPE;
        if (nameIdx >= 0 && typeIdx >= 0) {
          data.slice(1).forEach(function(row) {
            var name = row[nameIdx];
            if (!name) return;
            var key = normalizeAccountNameKey_(name);
            if (!key) return;
            map[key] = String(row[typeIdx] || '').trim();
          });
        }
      }
    }
    try {
      CacheService.getScriptCache().put(
        ACCOUNT_TYPE_MAP_CACHE_KEY,
        JSON.stringify(map),
        ACCOUNT_TYPE_MAP_TTL_SEC
      );
    } catch (cacheWriteErr) {
      Logger.log('loadAccountTypeByNameMap_ cache write warning: ' + cacheWriteErr.toString());
    }
  } catch (error) {
    Logger.log('loadAccountTypeByNameMap_ error: ' + error.toString());
  }
  return map;
}

function getAccountTypeForName_(accountName, typeMap) {
  typeMap = typeMap || loadAccountTypeByNameMap_();
  return typeMap[normalizeAccountNameKey_(accountName)] || '';
}

/** Deposits to Asset accounts are cash movements, not sales payment channels. */
function isDepositAccountExcludedFromSalesCalc_(accountName, typeMap) {
  var accountType = getAccountTypeForName_(accountName, typeMap);
  if (!accountType) return false;
  return accountType.toLowerCase() === 'asset';
}

function getInvestmentAccounts() {
  try {
    Logger.log("=== Getting Investment Accounts ===");
    return getAccountsByType('Investment');
  } catch (error) {
    Logger.log("Error getting investment accounts: " + error.toString());
    return [];
  }
}

function getAccountsByType(type) {
  try {
    Logger.log("Getting accounts by type: " + type);
    
    // Check if SHEETS is accessible
    if (typeof SHEETS === 'undefined') {
      Logger.log("SHEETS configuration not accessible");
      return [];
    }
    
    // Get spreadsheet and sheet
    const spreadsheet = getGlobalSpreadsheet();
    if (!spreadsheet) {
      Logger.log("Unable to access spreadsheet");
      return [];
    }
    
    const sheet = spreadsheet.getSheetByName(SHEETS.ACCOUNTS);
    if (!sheet) {
      Logger.log("Accounts sheet not found");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log("Retrieved " + data.length + " rows from Accounts sheet");
    
    if (data.length < 2) {
      Logger.log("No data rows found (only headers or empty sheet)");
      return [];
    }
    
    const headers = data.shift();
    Logger.log("Headers: " + JSON.stringify(headers));

    const cols = buildColumnMap(headers, 'ACCOUNTS');
    const typeIdx = cols.TYPE;
    const nameIdx = cols.NAME;
    
    Logger.log("Column indices - Type: " + typeIdx + ", Name: " + nameIdx);
    
    if (typeIdx === -1 || nameIdx === -1) {
      Logger.log("Required columns not found in Accounts sheet");
      return [];
    }
    
    const accounts = data
      .filter(row => row[typeIdx] === type)
      .map(row => row[nameIdx])
      .filter(Boolean);
      
    Logger.log("Returning " + accounts.length + " accounts of type '" + type + "': " + JSON.stringify(accounts));
    
    return accounts;
    
  } catch (error) {
    Logger.log("Error getting accounts by type: " + error.toString());
    return [];
  }
}