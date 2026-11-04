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
