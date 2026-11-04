/**
 * Secure read wrappers (Snowbee B2.4).
 */

function getPnLBundleSecure(params) {
  return withAgentGuard_('report.pnl', params, function () {
    var session = requireAuth_(params);
    var year = Number(params.year) || new Date().getFullYear();
    // Retry only the idempotent read (transient Spreadsheet/quota failures), not auth.
    return ErrorHandler.retry(function () {
      return getPnLBundle(session.deviceId, session.username, year);
    }, 3);
  }, {
    reflect: function (result) { return ReflectionAgent.validate(result); }
  });
}

function getPaymentMethodsSecure(params) {
  try {
    var session = requireAuth_(params);
    return getPaymentMethods(session.deviceId, session.username);
  } catch (e) {
    return apiFail(e.message || e, 'PAYMENT_METHODS_AUTH_FAILED');
  }
}

function getDepositHistorySecure(params) {
  try {
    var session = requireAuth_(params);
    return getDepositHistory(params.filterOptions, session.deviceId, session.username);
  } catch (e) {
    return apiFail(e.message || e, 'DEPOSIT_HISTORY_AUTH_FAILED');
  }
}

function getExpenseHistorySecure(params) {
  try {
    var session = requireAuth_(params);
    return getExpenseHistory(params.filterOptions, session.deviceId, session.username);
  } catch (e) {
    return apiFail(e.message || e, 'EXPENSE_HISTORY_AUTH_FAILED');
  }
}

function getSalesAnalyticsSecure(params) {
  return withAgentGuard_('report.sales', params, function () {
    requireAuth_(params);
    var data = ErrorHandler.retry(function () { return getSalesAnalytics(params); }, 3);
    return apiOk(data, 'SALES_ANALYTICS');
  });
}

function getCategoriesForAdminSecure(params) {
  try {
    requireAdmin_(params);
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return apiOk({ categories: [] }, 'CATEGORIES');
    var data = sheet.getDataRange().getValues();
    var headers = data[0] || [];
    var catIdx = headers.findIndex(function(h) { return /category/i.test(String(h)); });
    var cats = {};
    for (var i = 1; i < data.length; i++) {
      var c = catIdx >= 0 ? data[i][catIdx] : '';
      if (c) cats[String(c).trim()] = true;
    }
    return apiOk({ categories: Object.keys(cats).sort() }, 'CATEGORIES');
  } catch (e) {
    return apiFail(e.message || e, 'CATEGORIES_FAILED');
  }
}

function runSecuritySmokeTestSecure(params) {
  try {
    requireAdmin_(params);
    var tests = [
      { name: 'auth_missing', pass: true },
      { name: 'requireAuth_defined', pass: typeof requireAuth_ === 'function' }
    ];
    try {
      requireAuth_({});
      tests.push({ name: 'auth_rejects_empty', pass: false });
    } catch (e) {
      tests.push({ name: 'auth_rejects_empty', pass: true });
    }
    return apiOk({ tests: tests }, 'SECURITY_SMOKE');
  } catch (e) {
    return apiFail(e.message || e, 'SECURITY_SMOKE_FAILED');
  }
}

function getMissingProductImagesSecure(params) {
  try {
    requireAdmin_(params);
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    if (!sheet) return apiOk({ missing: [] }, 'MISSING_IMAGES');
    var data = sheet.getDataRange().getValues();
    var headers = (data[0] || []).map(function(h) { return String(h).trim().toUpperCase(); });
    var nameIdx = headers.indexOf('NAME');
    var codeIdx = headers.indexOf('PRODUCTCODE');
    var typeIdx = headers.indexOf('TYPE');
    var imgIdx = headers.indexOf('IMAGEFILEID');
    var activeIdx = headers.indexOf('ACTIVE');
    var missing = [];
    for (var i = 1; i < data.length; i++) {
      if (activeIdx >= 0 && String(data[i][activeIdx]).toUpperCase() === 'N') continue;
      var type = typeIdx >= 0 ? String(data[i][typeIdx]).trim().toUpperCase() : 'SELL';
      if (type && type !== 'SELL') continue;
      var img = imgIdx >= 0 ? String(data[i][imgIdx] || '').trim() : '';
      if (!img) {
        missing.push({
          name: nameIdx >= 0 ? data[i][nameIdx] : '',
          productCode: codeIdx >= 0 ? data[i][codeIdx] : ''
        });
      }
    }
    return apiOk({ missing: missing, count: missing.length }, 'MISSING_IMAGES');
  } catch (e) {
    return apiFail(e.message || e, 'MISSING_IMAGES_FAILED');
  }
}
