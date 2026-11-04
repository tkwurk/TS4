/**
 * reportingWorker.gs — Reporting / Analytics worker agent.
 *
 * Additive namespace that DELEGATES to the existing analytics functions.
 *
 * Depends on existing globals: getPnLBundle (pnlBundle.js), getCashSummary
 * (cashManagement.js), getSalesSummaryByDate (salesReports.js),
 * getDailyClosingBundle (inventoryConsumption.js), Repo (repo.gs), apiOk/apiFail.
 */

// Copilot: Implement a Reporting Worker agent for reports and summaries.
// Given a report type, delegate to the existing analytics functions
// (getPnLBundle, getCashSummary, getSalesSummaryByDate, getDailyClosingBundle)
// and return the report object. The result is sent to the Reflection Agent for
// validation before it reaches the Orchestrator.
var ReportingWorker = {

  /**
   * Produces a report by type, delegating to existing analytics.
   * @param {string} type one of 'pnl' | 'cash' | 'sales' | 'dailyClosing'
   * @param {Object} params report-specific params (e.g. { year, date })
   * @param {Object} ctx { deviceId, username, outlet, requestId }
   * @return {Object} apiOk wrapping the report, or apiFail.
   */
  generateReport: function (type, params, ctx) {
    var c = ctx || {};
    var p = params || {};
    var report;
    switch (String(type || 'dailyClosing')) {
      case 'pnl':
        if (typeof getPnLBundle !== 'function') return apiFail('getPnLBundle missing', 'REPORT_DELEGATE_MISSING');
        report = getPnLBundle(c.deviceId, c.username, p.year);
        break;
      case 'cash':
        if (typeof getCashSummary !== 'function') return apiFail('getCashSummary missing', 'REPORT_DELEGATE_MISSING');
        report = getCashSummary(c.deviceId, c.username);
        break;
      case 'sales':
        if (typeof getSalesSummaryByDate !== 'function') return apiFail('getSalesSummaryByDate missing', 'REPORT_DELEGATE_MISSING');
        report = getSalesSummaryByDate(params);
        break;
      case 'salesSummary':
        report = ReportingWorker._salesSummaryFromRepo(p, c);
        break;
      case 'recentOrders':
        report = ReportingWorker._recentOrdersFromRepo(p, c);
        break;
      case 'dailyClosing':
      default:
        if (typeof getDailyClosingBundle !== 'function') return apiFail('getDailyClosingBundle missing', 'REPORT_DELEGATE_MISSING');
        report = getDailyClosingBundle(c.deviceId, c.username);
        break;
    }
    return apiOk({ type: String(type || 'dailyClosing'), report: report }, 'REPORT_GENERATED');
  },

  /**
   * Repo-backed sales summary (Seam 2): aggregates the Orders sheet through the
   * Data Access Layer instead of SpreadsheetApp. Optional outlet filter.
   * @return {Object} { source, outlet, lineCount, totalSales, totalQty, byPayment }
   */
  _salesSummaryFromRepo: function (params, ctx) {
    var rows = Repo.readRows('ORDERS');
    var outlet = (params && params.outlet) || (ctx && ctx.outlet) || '';
    var want = String(outlet).trim().toLowerCase();
    var filtered = outlet
      ? rows.filter(function (r) {
          return String(ReportingWorker._field_(r, 'Outlet') || '').trim().toLowerCase() === want;
        })
      : rows;

    var totalSales = 0, totalQty = 0, byPayment = {};
    filtered.forEach(function (r) {
      totalSales += Number(ReportingWorker._field_(r, 'Total')) || 0;
      totalQty += Number(ReportingWorker._field_(r, 'Quantity', 'Qty')) || 0;
      var pay = String(ReportingWorker._field_(r, 'Payment') || 'unknown').trim() || 'unknown';
      byPayment[pay] = (byPayment[pay] || 0) + (Number(ReportingWorker._field_(r, 'Total')) || 0);
    });

    return {
      source: 'repo',
      outlet: outlet || null,
      lineCount: filtered.length,
      totalSales: totalSales,
      totalQty: totalQty,
      byPayment: byPayment
    };
  },

  /** Case-insensitive header lookup; pass one or more candidate names. */
  _field_: function (obj) {
    var keys = Object.keys(obj);
    for (var a = 1; a < arguments.length; a++) {
      var want = String(arguments[a]).toUpperCase();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i].toUpperCase() === want) return obj[keys[i]];
      }
    }
    return undefined;
  },

  /**
   * Repo-backed recent orders (Seam 2): returns the most recently appended order
   * lines from the Orders sheet (newest first). Optional outlet filter + limit.
   * @return {Object} { source, outlet, count, orders[] }
   */
  _recentOrdersFromRepo: function (params, ctx) {
    var rows = Repo.readRows('ORDERS');
    var outlet = (params && params.outlet) || (ctx && ctx.outlet) || '';
    var want = String(outlet).trim().toLowerCase();
    var filtered = outlet
      ? rows.filter(function (r) {
          return String(ReportingWorker._field_(r, 'Outlet') || '').trim().toLowerCase() === want;
        })
      : rows;

    var limit = Number(params && params.limit) || 25;
    var recent = filtered.slice(Math.max(0, filtered.length - limit)).reverse();
    var orders = recent.map(function (r) {
      return {
        timestamp: ReportingWorker._field_(r, 'Timestamp'),
        orderNumber: ReportingWorker._field_(r, 'Order Number', 'OrderNumber'),
        product: ReportingWorker._field_(r, 'Product Name', 'ProductName', 'Name'),
        qty: ReportingWorker._field_(r, 'Quantity', 'Qty'),
        price: ReportingWorker._field_(r, 'Price'),
        total: ReportingWorker._field_(r, 'Total'),
        payment: ReportingWorker._field_(r, 'Payment'),
        outlet: ReportingWorker._field_(r, 'Outlet'),
        username: ReportingWorker._field_(r, 'Username')
      };
    });
    return { source: 'repo', outlet: outlet || null, count: orders.length, orders: orders };
  }
};
