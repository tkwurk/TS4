# Orchestrator–Worker Agent Architecture — Design

> **Status:** Design only. No code has been changed. This document describes how
> to reorganize the POS into an orchestrator–worker–reflection architecture in a
> way that is **compatible with the Google Apps Script (V8) runtime** and does
> **not** break the live app.

---

## 1. Runtime constraints (why the classic module split cannot be literal)

The POS runs on **Google Apps Script V8** (`appsscript.json → runtimeVersion: V8`)
as an `HtmlService` web app. Two constraints shape the whole design:

1. **No module system.** GAS does not support ES modules (`import`/`export`) or
   CommonJS (`require`/`module.exports`). Every `.js`/`.gs` file shares **one
   global scope**. Introducing `import`/`export` is a parse error and the
   deployment fails.
2. **`google.script.run` can only call top-level global functions.** The HTML
   frontend invokes server functions by bare name — e.g.
   `google.script.run.saveOrderSecure(...)`, `getPnLBundle(...)`,
   `recordPurchaseSecure(...)`. Functions nested inside objects/modules are
   **not** reachable from the client.

**Consequence:** the architecture is implemented **additively**. Existing global
functions remain the real `google.script.run` entry points. Workers are
**namespace objects that delegate** to those functions. The orchestrator is a
new coordination layer on top — it changes *organization and flow*, never the
business logic.

---

## 2. Subsystem → Agent mapping (all five)

Each worker **wraps** existing functions; nothing is rewritten or moved out of
global scope.

### 2.1 Sales / Checkout → `salesWorker.gs`
Source: `orders.js`, `ui_checkout.html`, `ui_orders.html`.
Key existing functions it will delegate to:
- `saveOrder(orderJson)` / `saveOrderSecure(params)` — persist an order
- `validateOrderCart_(order)` — cart validation
- `buildOrderRows_` / `buildOrderItemRow_` — row construction
- `generateOrderNumber()`
- `buildAutoDepositRows_` / `shouldAutoDepositForMethod_` — auto-deposit
- `saveMissedOrdersSecure(params)`

### 2.2 Inventory / Stock → `inventoryWorker.gs`
Source: `inventoryConsumption.js`, `stockHistory.js`, `purchases.js`, `bom.js`.
Key functions:
- `recordInventoryConsumption(...)` / `recordInventoryConsumptionSecure(params)`
- `getTodaysConsumptionSummary(...)`, `getDailyClosingBundle(...)`
- `ensureStockHistoryCurrentForOutlet_`, `advanceStockHistoryForOutlet_`,
  `getStockTotalsForOutlet_`, `computeStockDeltasForDate_` (stock ledger)
- `recordPurchase(...)` / `recordPurchaseSecure(params)` (restocking)

### 2.3 Reporting / Analytics → `reportingWorker.gs`
Source: `pnlBundle.js`, `profitLoss.js`, `cashManagement.js`,
`monthlySalesReport.js`, `salesReports.js`, `usage.js`, `hourlyReport.js`.
Key functions:
- `getPnLBundle(...)`, `getProfitLossReport(...)`, `getPnLBundleSecure(params)`
- `getCashSummary(...)`, `getMonthlyCashSummary(...)`, `getDailyCashSummary(...)`
- `getSalesSummaryData(...)`, `getSalesSummaryByDate(params)`, `getOrderLinesByDate(params)`
- `getMonthlyUsageReport(...)`, monthly sales report builders

### 2.4 Validation / Compliance (Reflection) → `reflectionAgent.gs`
Source: `accessControl.js`, `secureReads.js`, `config.js`
(`SCHEMA_VALIDATION`, `VALIDATED_SHEETS`), `serverSecurity.js`, `unifiedAuth.js`.
Responsibilities:
- **AuthZ/AuthN gates:** `requireAuth_`, `requireAdmin_`, `requirePermission_`
- **Schema/consistency checks:** validate report output reconciles
  (e.g. `profit == sales − purchases − expenses`, cash-in-hand formula),
  detect negative stock, verify sheet schemas via `SCHEMA_VALIDATION`
- Returns a structured verdict `{ valid, issues[] }` to the orchestrator

### 2.5 Error Handling / Logging → `errorHandler.gs`
Source: `utils.js` (`apiOk`/`apiFail`), `auditLog.js` (`writeAuditLog_`),
`snowbeeNotify.js` (alerting), GAS `Logger` + Stackdriver.
Responsibilities:
- `wrap(fn, ctx)` — try/catch envelope returning `apiFail` on error
- `log(error, ctx)` — audit + Stackdriver
- `retry(fn, attempts)` — transient-failure retry (lock/quota)
- Alert routing via Snowbee for critical failures

---

## 3. Target file layout (new, additive)

```
orchestrator.gs      // central routing + workflow state/log
salesWorker.gs       // SalesWorker namespace
inventoryWorker.gs   // InventoryWorker namespace
reportingWorker.gs   // ReportingWorker namespace
reflectionAgent.gs   // ReflectionAgent namespace
errorHandler.gs      // ErrorHandler namespace
```

Existing files are untouched except that their global entry points may
**optionally** delegate into the orchestrator later (Phase 2). Workers are plain
global objects:

```js
// salesWorker.gs
var SalesWorker = {
  handleCheckout: function (payload, ctx) {
    // delegate to the existing, unchanged business logic
    return saveOrderSecure(payload);
  }
};
```

> `.gs` and `.js` are interchangeable in GAS; new files use `.gs` by convention
> for server agents. clasp pushes both (`scriptExtensions: [".js", ".gs"]`).

---

## 4. Worker interfaces

| Agent | Interface | Delegates to | Returns |
|---|---|---|---|
| `SalesWorker` | `handleCheckout(payload, ctx)` | `saveOrderSecure` / `saveOrder` | `{ success, orderNumbers, lines }` |
| `InventoryWorker` | `updateStock(payload, ctx)` | `recordInventoryConsumptionSecure`, stock-history advance | `{ success, deltas }` |
| `InventoryWorker` | `restock(payload, ctx)` | `recordPurchaseSecure` | `{ success, purchaseId }` |
| `ReportingWorker` | `generateReport(type, params, ctx)` | `getPnLBundle`, `getCashSummary`, `getSalesSummaryByDate`, … | report object |
| `ReflectionAgent` | `validate(report, ctx)` | reconciliation + `requirePermission_` + schema checks | `{ valid, issues[] }` |
| `ErrorHandler` | `wrap(fn, ctx)` / `log(err, ctx)` / `retry(fn, n)` | `apiFail`, `writeAuditLog_`, Snowbee | envelope / void |
| `Orchestrator` | `routeEvent(eventType, payload, ctx)` | the above | final validated output |

All methods take an optional `ctx` (`{ username, outlet, deviceId, requestId }`)
so the reflection and error layers have audit context.

---

## 5. Dependency wiring / event flow

> GAS is **synchronous** per request. "Triggers" below are **direct function
> calls** inside one server invocation, not async events. (A decoupled variant
> using an events sheet + time-driven triggers is noted in §8.)

```
Client (ui_checkout.html)
        │  google.script.run.processCheckoutEvent(payload)
        ▼
┌──────────────────────────────────────────────────────────────┐
│ Orchestrator.routeEvent('checkout.completed', payload, ctx)   │
│  (logs routing decision + workflow state)                     │
└───┬──────────────────────────────────────────────────────────┘
    │ 1. ErrorHandler.wrap →
    ▼
  SalesWorker.handleCheckout(payload)        → saveOrderSecure
    │ on success, emits stock-consumption intent
    ▼
  InventoryWorker.updateStock(order)         → recordInventoryConsumptionSecure
    │ on success, emits reporting intent
    ▼
  ReportingWorker.generateReport('daily')    → getDailyClosingBundle / getPnLBundle
    │ passes report to
    ▼
  ReflectionAgent.validate(report)           → reconcile + compliance
    │ { valid, issues[] }
    ▼
  Orchestrator  ── returns final output (or ErrorHandler.fail) to client
```

**Wiring rules (from the request):**
1. `SalesWorker` → triggers `InventoryWorker` after checkout.
2. `InventoryWorker` → triggers `ReportingWorker` after stock update.
3. `ReportingWorker` → sends results to `ReflectionAgent`.
4. `ReflectionAgent` → validates and returns final output to `Orchestrator`.

Every hop is wrapped by `ErrorHandler.wrap`, so a failure short-circuits with a
logged `apiFail` and (for critical events) a Snowbee alert.

---

## 6. Orchestrator routing table

```js
// orchestrator.gs (sketch)
var Orchestrator = {
  ROUTES: {
    'checkout.completed':  ['sales', 'inventory', 'reporting', 'reflection'],
    'stock.updated':       ['inventory', 'reporting', 'reflection'],
    'stock.report.requested': ['stockReport'],
    'stock.history.status.requested': ['stockHistoryStatus'],
    'stock.history.schedule.requested': ['stockHistorySchedule'],
    'report.requested':    ['reporting', 'reflection'],
    'purchase.recorded':   ['inventory', 'reporting', 'reflection']
  },
  routeEvent: function (eventType, payload, ctx) {
    WorkflowLog.record(eventType, 'START', ctx);
    // iterate the pipeline for eventType, passing each step's output forward,
    // wrapping each with ErrorHandler.wrap(...)
  }
};
```

- **Workflow state** is logged to a new `WorkflowLog` sheet (or reuses
  `AuditLog`) with `requestId, eventType, step, status, ts`.
- Routing decisions are logged so a run can be traced end to end.

---

## 7. Preserving existing behavior (non-negotiable)

- All current global entry points (`saveOrderSecure`, `getPnLBundle`,
  `recordPurchaseSecure`, …) **remain callable by `google.script.run`**.
- Workers **call** those functions; they do not replace or move them.
- A new thin entry point (e.g. `processCheckoutEvent(params)`) exposes the
  orchestrated pipeline to the client **without** removing the direct calls, so
  rollout is incremental and reversible.

---

## 8. Optional: true event decoupling (future)

If asynchronous decoupling is ever wanted (GAS cannot do in-process async):
- Workers append an **event row** to an `Events` sheet instead of calling the
  next worker directly.
- A **time-driven trigger** runs `Orchestrator.drainEvents()` to process queued
  events (heavy reports, reflection) off the request path.
This is a later enhancement, not part of the first implementation.

---

## 9. Phased migration plan (low risk)

| Phase | Deliverable | Risk |
|---|---|---|
| 0 | Scaffolding: 6 new files, namespaces + `// Copilot:` agent comments, workers delegate 1:1 | None (additive, nothing wired) |
| 1 | Wire the **checkout chain** behind a new `processCheckoutEvent`; keep `saveOrderSecure` as-is | Low (opt-in) |
| 2 | Route reporting + reflection; add `WorkflowLog` | Low |
| 3 | Centralize error/retry via `ErrorHandler.wrap` at entry points | Medium (touches entry points) |
| 4 | (Optional) event-sheet decoupling | Medium |

Each phase is independently deployable and reversible.

---

## 10. Non-goals / explicit risks

- **Not** microservices, not async, not a message bus — single synchronous GAS runtime.
- **No** `import`/`export`/`require`.
- **No** relocation of `google.script.run` entry points into objects.
- Business logic is **unchanged**; only organization and flow are added.

---

## 11. Agent comments

Inline `// Copilot: …` comments for each worker's main function are defined in
[`AGENT_GUIDE.md`](./AGENT_GUIDE.md). During Phase 0 they are placed above each
worker method.
