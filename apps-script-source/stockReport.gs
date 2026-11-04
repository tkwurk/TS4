/**
 * stockReport.gs — Stock tab read for the PWA.
 *
 * Thin, additive wrapper over the existing stock-ledger logic
 * (getStockTotalsForOutlet_ in stockHistory.js) so the Stock screen can read
 * current totals over the RPC bridge. Read-only; joins product names from the
 * Products sheet. Authenticated via _authContext like the other *Secure reads.
 */
function getStockTotalsSecure(params) {
  try {
    var session = requireAuth_(params);
    var outlet = (params.outlet || session.outlet || '').toString().trim();
    if (typeof getStockTotalsForOutlet_ !== 'function') {
      return apiFail('Stock ledger not available', 'STOCK_UNAVAILABLE');
    }
    var totals = getStockTotalsForOutlet_(outlet, {}) || {};
    var purchased = totals.purchased || {};
    var consumed = totals.consumed || {};
    var sold = totals.sold || {};

    // Product name/unit/category lookup by code.
    var nameByCode = {};
    try {
      var rows = (typeof Repo !== 'undefined') ? Repo.readRows('PRODUCTS') : [];
      rows.forEach(function (r) {
        var code = '';
        var name = '', unit = '', cat = '';
        Object.keys(r).forEach(function (k) {
          var u = k.toUpperCase().replace(/\s+/g, '');
          if (u === 'PRODUCTCODE' || u === 'CODE') code = String(r[k] || '').trim();
          else if (u === 'NAME' || u === 'PRODUCTNAME') name = String(r[k] || '');
          else if (u === 'UNIT' || u === 'BUYUNIT') unit = unit || String(r[k] || '');
          else if (u === 'CATEGORY') cat = String(r[k] || '');
        });
        if (code) nameByCode[code.toUpperCase()] = { name: name, unit: unit, category: cat };
      });
    } catch (e) { /* names optional */ }

    var codes = {};
    [purchased, consumed, sold].forEach(function (m) { Object.keys(m).forEach(function (k) { codes[k] = true; }); });
    var items = Object.keys(codes).map(function (code) {
      var meta = nameByCode[String(code).toUpperCase()] || {};
      var p = Number(purchased[code]) || 0, c = Number(consumed[code]) || 0, s = Number(sold[code]) || 0;
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
      historyIncomplete: !!totals.historyIncomplete
    }, 'STOCK_OK');
  } catch (e) {
    return apiFail(e.message || e, 'STOCK_FAILED');
  }
}
