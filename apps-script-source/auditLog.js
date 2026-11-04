/**
 * Audit log (Snowbee B3).
 */

var AUDIT_HEADER = ['timestamp', 'username', 'action', 'entity_id', 'details', 'outlet'];

function ensureAuditLogSheet_() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.AUDIT_LOG);
    sheet.appendRow(AUDIT_HEADER);
  }
  return sheet;
}

function writeAuditLog_(username, action, entityId, details, outlet) {
  try {
    var sheet = ensureAuditLogSheet_();
    sheet.appendRow([
      new Date(),
      username || '',
      action || '',
      entityId || '',
      (details || '').toString().substring(0, 500),
      outlet || ''
    ]);
  } catch (e) {
    Logger.log('writeAuditLog_ failed: ' + e);
  }
}

function softDeleteFlag_(sheetName, rowNum, flagColumn, value) {
  var sheet = getGlobalSpreadsheet().getSheetByName(sheetName);
  if (!sheet || !rowNum) return false;
  sheet.getRange(rowNum, flagColumn).setValue(value || 'DELETED');
  return true;
}

/** Orders tab: manual cash drawer open (single-use PIN; deactivates after log). */
function logCashDrawerOpenSecure(params) {
  try {
    var session = requireAuth_(params);
    var orig = extractOriginalParams(params);
    var outlet = (orig.outlet || session.outlet || '').toString().trim();
    if (typeof resolveUserOutlet_ === 'function' && session.username) {
      outlet = resolveUserOutlet_(session.username, outlet);
    }
    var printerName = (orig.printerName || '').toString().trim();
    var result = (orig.result || '').toString().trim();
    var openedOk = result === 'ok';
    if (openedOk) {
      _finalizeDrawerOpen_(session.username, outlet);
    } else {
      _assertDrawerPendingOpen_(session.username, outlet);
    }
    writeAuditLog_(
      session.username,
      openedOk ? 'DRAWER_OPEN' : 'DRAWER_OPEN_FAIL',
      printerName || 'unknown',
      JSON.stringify({
        result: result || 'unknown',
        printer: printerName,
        clientAt: (orig.clientAt || '').toString(),
        deactivated: openedOk
      }),
      outlet
    );
    return sanitizeForClient({
      success: true,
      message: openedOk ? 'Drawer open logged' : 'Drawer open attempt logged',
      deactivated: openedOk
    });
  } catch (e) {
    Logger.log('logCashDrawerOpenSecure error: ' + e.toString());
    return { success: false, error: e.message || e.toString() };
  }
}
