/**
 * Legacy internal entry point. The PWA uses stock.report.requested through the
 * Orchestrator; this wrapper remains for existing Apps Script callers.
 */
function getStockTotalsSecure(params) {
  var ctx = Orchestrator.buildContext_(params, 'stock.report.requested');
  return InventoryWorker.getStockReport(params, ctx);
}
