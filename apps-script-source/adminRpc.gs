/**
 * adminRpc.gs — admin-only wrappers for legacy trigger functions used by the
 * PWA Admin panel. The underlying smartPicks.js functions take no auth
 * (they were only reachable from the admin HTML), so they must not be exposed
 * directly over the public RPC bridge.
 */
function getSmartPicksTriggerInfoSecure(params) {
  try {
    requireAdmin_(params);
    return apiOk(getSmartPicksTriggerInfo(), 'SMARTPICKS_INFO');
  } catch (e) {
    return apiFail(e.message || e, 'SMARTPICKS_INFO_FAILED');
  }
}

function scheduleSmartPicksTriggerSecure(params) {
  try {
    requireAdmin_(params);
    var hour = parseInt(params.hour, 10);
    if (isNaN(hour) || hour < 0 || hour > 23) hour = 0;
    return apiOk(scheduleSmartPicksTrigger(hour), 'SMARTPICKS_SCHEDULED');
  } catch (e) {
    return apiFail(e.message || e, 'SMARTPICKS_SCHEDULE_FAILED');
  }
}

function removeSmartPicksTriggerSecure(params) {
  try {
    requireAdmin_(params);
    return apiOk(removeSmartPicksTrigger(), 'SMARTPICKS_REMOVED');
  } catch (e) {
    return apiFail(e.message || e, 'SMARTPICKS_REMOVE_FAILED');
  }
}

function refreshSmartPicksSecure(params) {
  try {
    requireAdmin_(params);
    return apiOk(refreshSmartPicks(), 'SMARTPICKS_REFRESHED');
  } catch (e) {
    return apiFail(e.message || e, 'SMARTPICKS_REFRESH_FAILED');
  }
}

// ── Cooking stock alerts (cookingAlerts.gs) ─────────────────────────────────
function _cookAlertTriggerInstalled_() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'runCookingStockAlertSweep';
  });
}

function getCookingAlertStatusSecure(params) {
  try {
    requireAdmin_(params);
    var recipients = null;
    try {
      var r = _cookingAlertRecipients_((params._authContext && params._authContext.outlet) || '');
      recipients = r && r.to ? r.to : (Array.isArray(r) ? r : null);
    } catch (e) { /* optional */ }
    var quota = null;
    try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { /* optional */ }
    return apiOk({
      installed: _cookAlertTriggerInstalled_(),
      sweepMinutes: COOK_ALERT_SWEEP_MINUTES,
      reminderMinutes: Math.round(COOK_ALERT_REMINDER_MS / 60000),
      recipients: recipients,
      mailQuotaRemaining: quota
    }, 'COOK_ALERT_STATUS');
  } catch (e) {
    return apiFail(e.message || e, 'COOK_ALERT_STATUS_FAILED');
  }
}

function installCookingAlertTriggerSecure(params) {
  try {
    requireAdmin_(params);
    installCookingStockAlertTrigger();
    return apiOk({ installed: true }, 'COOK_ALERT_INSTALLED');
  } catch (e) {
    return apiFail(e.message || e, 'COOK_ALERT_INSTALL_FAILED');
  }
}

function removeCookingAlertTriggerSecure(params) {
  try {
    requireAdmin_(params);
    removeCookingStockAlertTrigger();
    return apiOk({ installed: false }, 'COOK_ALERT_REMOVED');
  } catch (e) {
    return apiFail(e.message || e, 'COOK_ALERT_REMOVE_FAILED');
  }
}

/** Run one sweep immediately (creates/refreshes cook requests and emails recipients). */
function runCookingAlertSweepNowSecure(params) {
  try {
    requireAdmin_(params);
    var summary = runCookingStockAlertSweep() || {};
    var quota = null;
    try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { /* optional */ }
    return apiOk({ ran: true, summary: summary, mailQuotaRemaining: quota }, 'COOK_ALERT_SWEPT');
  } catch (e) {
    return apiFail(e.message || e, 'COOK_ALERT_SWEEP_FAILED');
  }
}
