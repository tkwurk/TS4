/**
 * agentIntegration.gs — wires the agent cross-cutting concerns (error handling,
 * audit, workflow logging, non-blocking reflection) into the REAL entry points.
 *
 * Design guarantees (so production behavior is preserved):
 *   - On success, the wrapped function's result is returned UNCHANGED.
 *   - Reflection is non-blocking: it logs anomalies and never alters the result
 *     or throws.
 *   - WorkflowLog writes to its sheet only on failure/anomaly; normal traces go
 *     to Logger (so the checkout/purchase hot path stays fast).
 *   - retry is OPT-IN and intended for idempotent READS only (never for writes,
 *     which could double-commit).
 *
 * Depends on globals: ErrorHandler, WorkflowLog, Orchestrator (buildContext_),
 * apiFail.
 */

function withAgentGuard_(label, params, fn, opts) {
  opts = opts || {};
  var ctx = (typeof Orchestrator !== 'undefined' && Orchestrator.buildContext_)
    ? Orchestrator.buildContext_(params, label)
    : {
        username: (params && params._authContext && params._authContext.username) || '',
        outlet: (params && params._authContext && params._authContext.outlet) || '',
        requestId: '',
        eventType: label
      };

  var result;
  try {
    result = opts.retry ? ErrorHandler.retry(fn, opts.retry) : fn();
  } catch (e) {
    ErrorHandler.log(e, ctx, label);
    if (typeof WorkflowLog !== 'undefined') WorkflowLog.record(ctx.requestId, label, 'run', 'FAIL', ctx);
    return apiFail((e && e.message) || String(e), 'AGENT_' + String(label).toUpperCase().replace(/\W+/g, '_') + '_FAILED');
  }

  // Non-blocking reflection — logs anomalies, never changes the result.
  if (opts.reflect && result && result.success !== false) {
    try {
      var verdict = opts.reflect(result, ctx);
      if (verdict && verdict.valid === false) {
        ErrorHandler.log('Reflection issues: ' + JSON.stringify(verdict.issues || []), ctx, label + ':reflection');
        if (typeof WorkflowLog !== 'undefined') WorkflowLog.record(ctx.requestId, label, 'reflection', 'ISSUES', ctx);
      }
    } catch (ignore) { /* reflection must never break the transaction */ }
  }

  return result;
}
