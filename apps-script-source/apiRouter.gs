/**
 * apiRouter.gs — Seam 1: the single JSON API entry point for external clients.
 *
 * ARCHITECTURE (see /AGENTS.md): the Expo/web client talks to the backend through
 * ONE endpoint. doPost receives { eventType, payload } and DELEGATES to the
 * Orchestrator — it contains NO business logic (routing/transport only). The
 * existing doGet (HTML UI, code.js) and google.script.run entry points are
 * untouched; this is purely additive.
 *
 * Request body (POSTed as text/plain to avoid a CORS preflight):
 *   { "eventType": "checkout.completed",
 *     "payload": { ...order data..., "_authContext": { username, outlet, deviceId } } }
 *
 * Response: the standard apiOk / apiFail envelope as JSON.
 *
 * Deployment (Option C): Execute as = Me (owner); Who has access = Anyone (or
 * Anyone with Google account). Caller identity is carried in payload._authContext
 * and enforced by the existing secure delegates, exactly as the current UI does.
 */

function doPost(e) {
  try {
    var body = _parseApiRequest_(e);
    var eventType = body.eventType;
    var payload = body.payload || {};

    // Allow auth to arrive either inside payload._authContext or as body.auth.
    if (body.auth && !payload._authContext) payload._authContext = body.auth;

    if (!eventType) {
      return _apiJson_(apiFail('Missing "eventType" in request body.', 'BAD_REQUEST'));
    }
    if (eventType === 'health') {
      return _apiJson_(apiOk({ status: 'ok', time: new Date().toISOString() }, 'HEALTH_OK'));
    }
    if (eventType === 'auth.login') {
      return _apiJson_(_authEnvelope_(
        (typeof authenticateUserClean === 'function')
          ? authenticateUserClean(payload.username, payload.password, payload.deviceId)
          : { success: false, message: 'authenticateUserClean not available' }
      ));
    }
    if (eventType === 'auth.logout') {
      var who = payload._authContext || payload || {};
      return _apiJson_(_authEnvelope_(
        (typeof logoutUser === 'function')
          ? logoutUser(who.deviceId, who.username)
          : { success: false, message: 'logoutUser not available' }
      ));
    }

    var ctx = Orchestrator.buildContext_(payload, eventType);
    var result = Orchestrator.routeEvent(eventType, payload, ctx);
    return _apiJson_(result);
  } catch (err) {
    try { Logger.log('doPost error: ' + err); } catch (e2) { /* ignore */ }
    return _apiJson_(apiFail('Server error: ' + err, 'SERVER_ERROR'));
  }
}

/** Normalize the login/logout result objects into the apiOk/apiFail envelope. */
function _authEnvelope_(res) {
  if (res && res.success) {
    return apiOk({
      username: res.username,
      outlet: res.outlet,
      role: res.role,
      allowedTabs: res.allowedTabs,
      deviceId: res.deviceId
    }, 'AUTH_OK');
  }
  return apiFail((res && res.message) || 'Authentication failed', 'AUTH_FAILED');
}

/** Parse the POST body as JSON, tolerating text/plain and form-encoded delivery. */
function _parseApiRequest_(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      throw new Error('Request body is not valid JSON.');
    }
  }
  // Fallback: form parameters (?eventType=...&payload={...})
  if (e.parameter && e.parameter.eventType) {
    var payload = {};
    if (e.parameter.payload) {
      try { payload = JSON.parse(e.parameter.payload); } catch (err) { payload = {}; }
    }
    return { eventType: e.parameter.eventType, payload: payload };
  }
  return {};
}

/** Serialize an envelope as a JSON ContentService response. */
function _apiJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
