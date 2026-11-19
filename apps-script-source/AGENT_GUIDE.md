# AGENT_GUIDE.md

Agent definitions and the inline `// Copilot:` comments to place above each main
function/class when the orchestrator–worker architecture is implemented. See
[`ORCHESTRATOR_ARCHITECTURE.md`](./ORCHESTRATOR_ARCHITECTURE.md) for the full
design and the Google Apps Script constraints.

> **Placement rule:** put the matching comment block immediately **above** the
> named method in its worker file. Keep the existing business logic unchanged —
> the worker method only delegates to the current global functions.

---

## Orchestrator — `orchestrator.gs`

```js
// Copilot: Implement a central Orchestrator agent.
// Route incoming POS events (checkout.completed, stock.updated,
// stock.report.requested, stock.history.status.requested,
// stock.history.schedule.requested, report.requested, purchase.recorded) to the correct
// worker pipeline. Maintain workflow state,
// log every routing decision to WorkflowLog, wrap each hop in ErrorHandler, and
// return the Reflection-validated final output. Do NOT contain business logic.
var Orchestrator = {
  routeEvent: function (eventType, payload, ctx) { /* ... */ }
};
```

---

## Sales Worker — `salesWorker.gs`

```js
// Copilot: Implement a Sales Worker agent for checkout and transactions.
// Validate the cart, persist the order, and handle auto-deposit. Delegate to the
// existing saveOrderSecure/saveOrder logic unchanged. On success, return the
// order number(s) and lines so the Orchestrator can trigger the Inventory Worker.
var SalesWorker = {
  handleCheckout: function (payload, ctx) { /* delegates to saveOrderSecure */ }
};
```

---

## Inventory Worker — `inventoryWorker.gs`

```js
// Copilot: Implement an Inventory Worker agent for stock updates and restocking.
// After a checkout, apply consumption to the stock ledger; for purchases, apply
// restock. Delegate to recordInventoryConsumptionSecure, the stock-history
// advance functions, and recordPurchaseSecure unchanged. Serve stock reports
// from Repo-backed history plus the current day's tail scan.
var InventoryWorker = {
  updateStock: function (payload, ctx) { /* delegates to consumption/stock-history */ },
  getStockReport: function (payload, ctx) { /* reads stock history + today's deltas through Repo */ },
  restock:     function (payload, ctx) { /* delegates to recordPurchaseSecure */ }
};
```

---

## Reporting Worker — `reportingWorker.gs`

```js
// Copilot: Implement a Reporting Worker agent for reports and summaries.
// Given a report type, delegate to the existing analytics functions
// (getPnLBundle, getCashSummary, getSalesSummaryByDate, getDailyClosingBundle,
// getMonthlyUsageReport) and return the report object. Send the result to the
// Reflection Agent for validation before it reaches the Orchestrator.
var ReportingWorker = {
  generateReport: function (type, params, ctx) { /* delegates to analytics */ }
};
```

---

## Reflection Agent — `reflectionAgent.gs`

```js
// Copilot: Implement a Reflection Agent for validation and compliance.
// Verify permissions (requirePermission_), confirm report totals reconcile
// (profit == sales - purchases - expenses; cash-in-hand formula), validate sheet
// schemas via SCHEMA_VALIDATION, and flag anomalies (e.g. negative stock).
// Return { valid, issues[] } to the Orchestrator as the final gate.
var ReflectionAgent = {
  validate: function (report, ctx) { /* returns { valid, issues } */ }
};
```

---

```js
// Copilot: Implement a centralized Error Handler agent.
// Provide wrap(fn, ctx) to run any worker step inside a try/catch returning an
// apiFail envelope, log(error, ctx) to AuditLog + Stackdriver, and retry(fn, n)
// for transient Spreadsheet/quota failures. Route critical failures to Snowbee.
var ErrorHandler = {
  wrap:  function (fn, ctx) { /* try/catch -> apiOk/apiFail */ },
  log:   function (error, ctx) { /* writeAuditLog_ + Logger */ },
  retry: function (fn, attempts) { /* backoff on transient errors */ }
};
```

---

## Agent responsibility matrix

| Agent | Owns | Must NOT do |
|---|---|---|
| Orchestrator | Routing, workflow state, logging decisions | Business logic, data writes |
| Sales Worker | Checkout/transaction persistence | Stock, reporting, validation |
| Inventory Worker | Stock ledger + restock | Sales pricing, reporting |
| Reporting Worker | Report/summary generation | Writing transactional data |
| Reflection Agent | Validation, compliance, reconciliation | Mutating business data |
| Error Handler | Logging, retry, alerting | Business decisions |

---

## Context object (`ctx`)

All agent methods accept a shared context so logging/validation is consistent:

```js
ctx = { username, outlet, deviceId, requestId, eventType };
```

---

## Reminder (GAS compatibility)

- These are **namespace objects**, not ES modules. No `import`/`export`.
- The real `google.script.run` entry points stay as **global functions**; agents
  delegate to them.
