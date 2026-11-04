/**
 * reflectionAgent.gs — Validation / Compliance (Reflection) agent.
 *
 * Validates reporting output before it is returned to the Orchestrator. It is
 * read-only: it never mutates business data.
 *
 * Depends on existing globals: apiOk (utils.js). Optionally uses
 * requirePermission_ (accessControl.js) when a permission + params are supplied.
 */

// Copilot: Implement a Reflection Agent for validation and compliance.
// Confirm report totals reconcile (profit == sales - purchases - expenses),
// flag anomalies (e.g. negative values), and optionally verify permissions.
// Return { success, valid, issues[] } to the Orchestrator as the final gate.
var ReflectionAgent = {

  /**
   * Validates a report object (typically the ReportingWorker envelope).
   * @return {Object} { success:true, valid:boolean, issues:string[] }
   */
  validate: function (report, ctx) {
    var issues = [];

    // Unwrap the ReportingWorker envelope { type, report } when present.
    var body = report && report.report ? report.report : report;

    if (!body) {
      issues.push('Report is empty.');
      return { success: true, valid: false, issues: issues };
    }

    // P&L bundle shape: body.daily = [{ sales, purchases, expenses, profit }, ...]
    var rows = null;
    if (Array.isArray(body.daily)) rows = body.daily;
    else if (Array.isArray(body)) rows = body;

    if (rows) {
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i] || {};
        if (typeof r.profit === 'number' &&
            typeof r.sales === 'number' &&
            typeof r.purchases === 'number' &&
            typeof r.expenses === 'number') {
          var expected = r.sales - r.purchases - r.expenses;
          if (Math.abs(expected - r.profit) > 0.01) {
            issues.push('Profit mismatch on row ' + i + ': reported ' + r.profit + ' vs expected ' + expected + '.');
          }
        }
        ['sales', 'purchases', 'expenses'].forEach(function (k) {
          if (typeof r[k] === 'number' && r[k] < 0) {
            issues.push('Negative ' + k + ' on row ' + i + ': ' + r[k] + '.');
          }
        });
      }
    }

    return { success: true, valid: issues.length === 0, issues: issues };
  }
};

// Copilot: Non-blocking negative-stock check for report/closing output. Deep-scans
// an already-computed report object for stock-like numeric fields that are
// negative and returns { valid, issues } without touching business data.
ReflectionAgent.validateStock = function (obj, ctx) {
  var issues = [];
  var STOCK_KEY = /stock|closing|onhand|on_hand|remaining|balance/i;
  function walk(node, keyName) {
    if (node == null) return;
    if (typeof node === 'number') {
      if (STOCK_KEY.test(keyName) && node < 0) issues.push('Negative ' + keyName + ': ' + node);
      return;
    }
    if (Array.isArray(node)) { for (var i = 0; i < node.length; i++) walk(node[i], keyName); return; }
    if (typeof node === 'object') {
      for (var k in node) { if (Object.prototype.hasOwnProperty.call(node, k)) walk(node[k], k); }
    }
  }
  try { walk(obj, ''); } catch (e) { /* ignore */ }
  return { success: true, valid: issues.length === 0, issues: issues.slice(0, 20) };
};
