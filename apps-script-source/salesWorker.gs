/**
 * salesWorker.gs — Sales / Checkout worker agent.
 *
 * Additive namespace that DELEGATES to the existing, unchanged order logic. It
 * does not replace saveOrderSecure/saveOrder, which remain the real
 * google.script.run entry points.
 *
 * Depends on existing globals: saveOrderSecure (orders.js), apiOk/apiFail.
 */

// Copilot: Implement a Sales Worker agent for checkout and transactions.
// Validate the cart, persist the order, and handle auto-deposit. Delegate to the
// existing saveOrderSecure/saveOrder logic unchanged. On success, return the
// order number(s) and lines so the Orchestrator can trigger the Inventory Worker.
var SalesWorker = {

  /**
   * Persists a checkout. `payload` is the same params object the client already
   * sends to saveOrderSecure (order data + _authContext).
   * @return {Object} apiOk/apiFail envelope from the existing order logic.
   */
  handleCheckout: function (payload, ctx) {
    if (typeof saveOrderSecure !== 'function') {
      return apiFail('saveOrderSecure not available', 'SALES_DELEGATE_MISSING');
    }
    return saveOrderSecure(payload);
  }
};
