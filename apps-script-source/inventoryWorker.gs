/**
 * inventoryWorker.gs — Inventory / Stock worker agent.
 *
 * Additive namespace that DELEGATES to the existing stock-ledger and purchase
 * logic. Nothing is rewritten; existing entry points are untouched.
 *
 * Depends on existing globals: getStockTotalsForOutlet_ (stockHistory.js),
 * recordPurchaseSecure (purchases.js), isProductActive_ (products.js), Repo
 * (repo.gs), apiOk/apiFail.
 */

// Copilot: Implement an Inventory Worker agent for stock updates and restocking.
// After a checkout, apply consumption to the stock ledger; for purchases, apply
// restock. Delegate to the stock-history advance functions and
// recordPurchaseSecure unchanged. Return the applied result so the Orchestrator
// can trigger the Reporting Worker.
var InventoryWorker = {

  /**
   * Refreshes/advances the stock ledger for the context outlet after a sale.
   * Delegates to the existing stock-history maintenance.
   * @return {Object} apiOk with a stock-totals snapshot summary.
   */
  updateStock: function (payload, ctx) {
    var outlet = (ctx && ctx.outlet) || (payload && payload.outlet) || '';
    if (!outlet) return apiFail('No outlet in context', 'INVENTORY_NO_OUTLET');
    if (typeof getStockTotalsForOutlet_ !== 'function') {
      return apiFail('getStockTotalsForOutlet_ not available', 'INVENTORY_DELEGATE_MISSING');
    }
    var totals = getStockTotalsForOutlet_(outlet, {}) || {};
    var codes = totals && typeof totals === 'object' ? Object.keys(totals).length : 0;
    return apiOk({ outlet: outlet, itemCount: codes }, 'STOCK_UPDATED');
  },

  /**
   * Builds the Stock report from the Repo-backed ledger for the authenticated outlet.
   */
  getStockReport: function (payload, ctx) {
    var session = requirePermission_(payload, 'stockReport');
    var outlet = String(session.outlet || '').trim();
    var requestedOutlet = String((payload && payload.outlet) || '').trim();
    if (!outlet) return apiFail('No outlet is assigned to this account.', 'STOCK_NO_OUTLET');
    if (requestedOutlet && canonicalOutletKey_(requestedOutlet) !== canonicalOutletKey_(outlet)) {
      return apiFail('Requested outlet is not authorized.', 'STOCK_OUTLET_FORBIDDEN');
    }

    var totals = getStockTotalsForOutlet_(outlet, { requireCompleteHistory: true }) || {};
    if (totals.historyIncomplete) {
      return apiFail(
        'Stock history is not current yet. The scheduled stock-history job must complete before this report is available.',
        'STOCK_HISTORY_NOT_READY'
      );
    }

    var nameByCode = {};
    Repo.readRows('PRODUCTS').forEach(function (row) {
      var code = InventoryWorker._fieldCI_(row, 'Product Code') || InventoryWorker._fieldCI_(row, 'Code');
      if (!code) return;
      var key = normalizeProductCodeKey_(code);
      nameByCode[key] = {
        name: InventoryWorker._fieldCI_(row, 'Name', 'Product Name') || key,
        unit: InventoryWorker._fieldCI_(row, 'Buy Unit', 'BuyUnit', 'Unit') || '',
        category: InventoryWorker._fieldCI_(row, 'Category') || ''
      };
    });

    var purchased = totals.purchased || {};
    var consumed = totals.consumed || {};
    var sold = totals.sold || {};
    var codes = {};
    [purchased, consumed, sold].forEach(function (map) {
      Object.keys(map).forEach(function (code) { codes[code] = true; });
    });
    var items = Object.keys(codes).map(function (code) {
      var meta = nameByCode[normalizeProductCodeKey_(code)] || {};
      var p = Number(purchased[code]) || 0;
      var c = Number(consumed[code]) || 0;
      var s = Number(sold[code]) || 0;
      return {
        productCode: code,
        name: meta.name || code,
        unit: meta.unit || '',
        category: meta.category || '',
        purchased: p,
        consumed: c,
        sold: s,
        remaining: Math.round((p - c - s) * 1000) / 1000
      };
    });
    items.sort(function (a, b) { return a.remaining - b.remaining; });

    return apiOk({
      outlet: outlet,
      items: items,
      stockSource: totals.stockSource || '',
      historyAsOfDate: totals.historyAsOfDate || '',
      todayDate: totals.todayDate || '',
      historyIncomplete: false
    }, 'STOCK_OK');
  },

  // Copilot: Expose stock-history trigger status to authorized administrators.
  getStockHistoryJobStatus: function (payload, ctx) {
    requirePermission_(payload, 'adminOrders');
    if (typeof getNightlyStockHistoryTriggerInfo !== 'function') {
      return apiFail('Stock history trigger status is unavailable.', 'STOCK_HISTORY_STATUS_UNAVAILABLE');
    }
    return apiOk(getNightlyStockHistoryTriggerInfo(), 'STOCK_HISTORY_JOB_STATUS');
  },

  // Copilot: Schedule the stock-history batch job in the Apps Script timezone.
  scheduleStockHistoryJob: function (payload, ctx) {
    requirePermission_(payload, 'adminOrders');
    var rawHour = payload && payload.hour;
    var hour = Number(rawHour);
    if (rawHour === undefined || rawHour === null || String(rawHour).trim() === '' ||
        !isFinite(hour) || Math.floor(hour) !== hour || hour < 0 || hour > 23) {
      return apiFail('Schedule hour must be an integer from 0 to 23.', 'STOCK_HISTORY_INVALID_HOUR');
    }
    if (typeof scheduleNightlyStockHistoryTrigger !== 'function') {
      return apiFail('Stock history trigger scheduling is unavailable.', 'STOCK_HISTORY_SCHEDULE_UNAVAILABLE');
    }
    return apiOk(scheduleNightlyStockHistoryTrigger(hour), 'STOCK_HISTORY_JOB_SCHEDULED');
  },

  /**
   * Records a restock (purchase). `payload` is the params object the client
   * already sends to recordPurchaseSecure (purchase data + _authContext).
   * @return {Object} apiOk/apiFail envelope from the existing purchase logic.
   */
  restock: function (payload, ctx) {
    if (typeof recordPurchaseSecure !== 'function') {
      return apiFail('recordPurchaseSecure not available', 'RESTOCK_DELEGATE_MISSING');
    }
    return recordPurchaseSecure(payload);
  },

  /**
   * Lists active products for the context outlet — reads through the Data Access
   * Layer (Repo), NOT SpreadsheetApp, so this path is adapter-agnostic and
   * database-swappable (see /AGENTS.md, DATA_ACCESS_LAYER.md). Reuses the
   * existing `isProductActive_` rule and mirrors getProducts' outlet matching.
   * @param {Object} ctx { outlet, username, deviceId, requestId }
   * @param {Object} [opts] { outlet, type } overrides (outlet defaults to ctx.outlet)
   * @return {Object} apiOk with { outlet, count, products[] }.
   */
  listProducts: function (ctx, opts) {
    var c = ctx || {};
    var o = opts || {};
    var rows;
    try {
      rows = Repo.readRows('PRODUCTS');
    } catch (e) {
      return apiFail('Failed to read products: ' + e, 'PRODUCTS_READ_FAILED');
    }

    var outlet = (o.outlet != null ? o.outlet : c.outlet) || '';
    var typeFilter = (o.type || '').toString().trim();

    var list = rows.filter(function (p) {
      return (typeof isProductActive_ === 'function')
        ? isProductActive_(InventoryWorker._fieldCI_(p, 'ACTIVE'))
        : true;
    });

    if (outlet) {
      list = list.filter(function (p) { return InventoryWorker._outletMatches_(p, outlet); });
    }

    if (typeFilter) {
      var want = typeFilter.toLowerCase();
      list = list.filter(function (p) {
        var t = (InventoryWorker._fieldCI_(p, 'TYPE') || 'Sell').toString().trim().toLowerCase();
        return t === want;
      });
    }

    return apiOk({ outlet: outlet || null, count: list.length, products: list }, 'PRODUCTS_LISTED');
  },

  /** Case-insensitive lookup of a header-keyed field value with aliases. */
  _fieldCI_: function (obj) {
    var keys = Object.keys(obj || {});
    for (var n = 1; n < arguments.length; n++) {
      var up = String(arguments[n] || '').trim().toUpperCase();
      for (var i = 0; up && i < keys.length; i++) {
        if (keys[i].toString().trim().toUpperCase() !== up) continue;
        var value = obj[keys[i]];
        if (value !== '' && value !== null && value !== undefined) return value;
      }
    }
    return undefined;
  },

  /** Mirrors getProducts: empty OUTLET or 'ALL' includes; else comma-list membership. */
  _outletMatches_: function (p, userOutlet) {
    var productOutlet = (InventoryWorker._fieldCI_(p, 'OUTLET') || '').toString().trim();
    if (!productOutlet) return true;
    if (productOutlet.toLowerCase().indexOf('all') !== -1) return true;
    var listed = productOutlet.split(',').map(function (x) { return x.trim().toLowerCase(); });
    return listed.indexOf(userOutlet.toString().trim().toLowerCase()) !== -1;
  }
};

/**
 * inventoryListProductsTest — run from the Apps Script editor
 * (Run ▸ inventoryListProductsTest) to verify a worker reads live data through
 * Repo. Read-only. Logs the active-product count and returns a status string.
 */
function inventoryListProductsTest() {
  var res = InventoryWorker.listProducts({ outlet: '' }, {});
  var msg = res && res.success
    ? ('InventoryWorker.listProducts OK via Repo — products: ' + res.data.count)
    : ('InventoryWorker.listProducts FAILED: ' + (res && res.error));
  Logger.log(msg);
  return msg;
}
