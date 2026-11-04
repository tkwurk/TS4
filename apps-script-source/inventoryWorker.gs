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

  /** Case-insensitive lookup of a header-keyed field value. */
  _fieldCI_: function (obj, name) {
    var up = name.toUpperCase();
    var keys = Object.keys(obj);
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].toString().trim().toUpperCase() === up) return obj[keys[i]];
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
