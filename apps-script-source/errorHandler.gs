/**
 * errorHandler.gs — centralized Error Handling / Logging agent.
 *
 * Additive layer for the orchestrator–worker architecture. See
 * ORCHESTRATOR_ARCHITECTURE.md and AGENT_GUIDE.md. This file introduces no new
 * google.script.run entry points; it is used by the other agents.
 *
 * Depends on existing globals: apiOk, apiFail (utils.js), writeAuditLog_
 * (auditLog.js), Logger/Utilities (GAS).
 */

// Copilot: Implement a centralized Error Handler agent.
// Provide wrap(label, fn, ctx) to run any worker step inside a try/catch returning
// an apiFail envelope, log(error, ctx, label) to AuditLog + Stackdriver, and
// retry(fn, attempts) for transient Spreadsheet/quota failures.
var ErrorHandler = {

  /** Runs fn() inside a guarded envelope; returns fn()'s result or apiFail. */
  wrap: function (label, fn, ctx) {
    try {
      return fn();
    } catch (e) {
      ErrorHandler.log(e, ctx, label);
      return apiFail((e && e.message) || String(e), 'AGENT_' + String(label || 'STEP').toUpperCase() + '_FAILED');
    }
  },

  /** Audit + Stackdriver logging with workflow context. */
  log: function (error, ctx, label) {
    var c = ctx || {};
    var msg = '[AGENT_ERROR:' + (label || 'UNKNOWN') + '] ' + ((error && error.stack) || error);
    try { Logger.log(msg); } catch (e) { /* ignore */ }
    try {
      if (typeof writeAuditLog_ === 'function') {
        writeAuditLog_(c.username || 'system', 'AGENT_ERROR:' + (label || 'UNKNOWN'),
                       c.requestId || '', String((error && error.message) || error), c.outlet || '');
      }
    } catch (e) { /* never let logging throw */ }
  },

  /** Retries fn() up to `attempts` times with linear backoff on transient errors. */
  retry: function (fn, attempts) {
    var max = attempts || 3;
    var lastErr = null;
    for (var i = 0; i < max; i++) {
      try {
        return fn();
      } catch (e) {
        lastErr = e;
        try { Utilities.sleep(200 * (i + 1)); } catch (se) { /* ignore */ }
      }
    }
    throw lastErr;
  }
};
