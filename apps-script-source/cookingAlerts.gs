/**
 * cookingAlerts.gs — Automatic "please cook more" notifications.
 *
 * The app (not the cashier) watches cooked-stock levels. When a cooking product
 * reaches or drops below its recook threshold, it emails the kitchen team
 * (People.position = Chef) and the admin team (People.position = Manager /
 * Director) and keeps reminding them every 30 minutes until a cooking batch is
 * logged (which auto-resolves the alert in kitchenCooking.js).
 *
 * Reuses existing globals: getOutlets, getCookingProductsFromSheet,
 * _buildCookingStockSummary, ensureKitchenCookRequestsHeaders, _kcHeaders,
 * _readSheetAllRows_, _findOpenCookRequestRowIndex_, _isOpenCookRequestStatus_,
 * ensurePeopleSheet_, personFromRow_, _kitchenOutletMatches_,
 * formatTimestampForSheet, _parseSheetDate, _appTimeZone_.
 */

var COOK_ALERT_REMINDER_MS = 30 * 60 * 1000; // remind every 30 minutes
var COOK_ALERT_SWEEP_MINUTES = 5;            // how often the trigger checks stock

/** Time-trigger entry point: check every outlet and alert/remind as needed. */
function runCookingStockAlertSweep() {
  try {
    var keys = [];
    var outlets = (typeof getOutlets === 'function') ? getOutlets() : [];
    if (outlets && outlets.length) {
      outlets.forEach(function (o) {
        var k = (o.name || o.id || '').toString().trim();
        if (k && keys.indexOf(k) === -1) keys.push(k);
      });
    }
    if (!keys.length) keys = ['']; // single-outlet / unconfigured fallback

    keys.forEach(function (outlet) {
      try {
        _sweepCookingAlertsForOutlet_(outlet);
      } catch (e) {
        Logger.log('cook alert sweep failed for outlet "' + outlet + '": ' + e.toString());
      }
    });
  } catch (e) {
    Logger.log('runCookingStockAlertSweep error: ' + e.toString());
  }
}

function _sweepCookingAlertsForOutlet_(outlet) {
  var products = getCookingProductsFromSheet(outlet);
  if (!products.length) return;

  var now = new Date();
  var summary = _buildCookingStockSummary(outlet, now, { skipCache: true, products: products });
  var below = summary.filter(function (i) { return i.needsRecook; });

  var sheet = ensureKitchenCookRequestsHeaders();
  var idx = _kcHeaders(sheet);
  var data = _readSheetAllRows_(sheet);
  var nowStr = formatTimestampForSheet(now);

  var belowCodes = {};
  var dueItems = [];
  var pendingStampRows = [];
  var newAlertCount = 0;

  below.forEach(function (item) {
    belowCodes[(item.productCode || '').toString().trim().toUpperCase()] = true;
    var existing = _findOpenCookRequestRowIndex_(data, idx, outlet, item.productCode);

    if (existing >= 1) {
      var sheetRow = existing + 1;
      var range = sheet.getRange(sheetRow, 1, 1, sheet.getLastColumn());
      var vals = range.getValues()[0];
      vals[idx.REMAINING] = item.remaining;
      vals[idx.SEVERITY] = item.stockStatus;
      vals[idx.UPDATEDAT] = nowStr;
      range.setValues([vals]);

      var last = _parseSheetDate(vals[idx.LASTNOTIFIEDAT]);
      var due = !last || (now.getTime() - last.getTime()) >= COOK_ALERT_REMINDER_MS;
      if (due) {
        if (!last) newAlertCount++;
        dueItems.push(item);
        pendingStampRows.push(sheetRow);
      }
    } else {
      var requestId = 'REQ-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
      var row = [];
      for (var c = 0; c < sheet.getLastColumn(); c++) row.push('');
      row[idx.REQUESTID] = requestId;
      row[idx.CREATEDAT] = nowStr;
      row[idx.OUTLET] = outlet;
      row[idx.PRODUCTCODE] = item.productCode;
      row[idx.PRODUCTNAME] = item.name;
      row[idx.REQUESTEDBY] = 'auto';
      row[idx.STATUS] = 'REQUESTED';
      row[idx.SEVERITY] = item.stockStatus;
      row[idx.REMAINING] = item.remaining;
      row[idx.UPDATEDAT] = nowStr;
      row[idx.SOURCE] = 'auto';
      row[idx.NOTIFYCOUNT] = 0;
      sheet.appendRow(row);
      data.push(row); // keep local scan in sync to avoid same-sweep duplicates

      newAlertCount++;
      dueItems.push(item);
      pendingStampRows.push(sheet.getLastRow());
    }
  });

  if (dueItems.length) {
    var recipients = _cookingAlertRecipients_(outlet);
    var isReminder = newAlertCount === 0;
    var sent = _sendCookAlertEmail_(outlet, dueItems, isReminder, recipients);
    if (sent) {
      pendingStampRows.forEach(function (sheetRow) {
        var range = sheet.getRange(sheetRow, 1, 1, sheet.getLastColumn());
        var vals = range.getValues()[0];
        vals[idx.LASTNOTIFIEDAT] = nowStr;
        vals[idx.NOTIFYCOUNT] = (Number(vals[idx.NOTIFYCOUNT]) || 0) + 1;
        range.setValues([vals]);
      });
    }
  }

  _closeRecoveredAutoAlerts_(sheet, idx, outlet, belowCodes, nowStr);
}

/** Auto-close alerts for products that are no longer below threshold. */
function _closeRecoveredAutoAlerts_(sheet, idx, outlet, belowCodes, nowStr) {
  var data = _readSheetAllRows_(sheet);
  var ou = (outlet || '').toString().trim().toUpperCase();
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!_isOpenCookRequestStatus_(row[idx.STATUS])) continue;
    if ((row[idx.SOURCE] || '').toString().trim().toLowerCase() !== 'auto') continue;
    if (ou && (row[idx.OUTLET] || '').toString().trim().toUpperCase() !== ou) continue;
    if (belowCodes[(row[idx.PRODUCTCODE] || '').toString().trim().toUpperCase()]) continue;
    var range = sheet.getRange(r + 1, 1, 1, sheet.getLastColumn());
    var vals = range.getValues()[0];
    vals[idx.STATUS] = 'DONE';
    vals[idx.UPDATEDAT] = nowStr;
    if (!vals[idx.HANDLEDBY]) vals[idx.HANDLEDBY] = 'auto-recovered';
    range.setValues([vals]);
  }
}

/**
 * Resolve alert recipients from the People sheet.
 * - Chefs (position/role = chef) and Managers (position/role = manager): scoped
 *   to the alert's outlet (or outlet-wide if their outlet is blank / ALL).
 * - Directors (position = director) and admins (role = admin): company-wide.
 */
function _cookingAlertRecipients_(outlet) {
  var out = { to: [], chefs: [], admins: [] };
  var seen = {};
  try {
    var sheet = ensurePeopleSheet_();
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var p = personFromRow_(data[r]);
      if (!p || !p.email) continue;

      var active = (p.active || 'Y').toString().trim().toUpperCase();
      if (['N', 'NO', 'FALSE', '0'].indexOf(active) !== -1) continue;

      var email = p.email.toString().trim();
      if (!email || email.indexOf('@') === -1) continue;

      var pos = (p.position || '').toString().trim().toLowerCase();
      var role = (p.role || '').toString().trim().toLowerCase();
      var isChef = pos.indexOf('chef') !== -1 || role === 'chef';
      var isManager = pos.indexOf('manager') !== -1 || role === 'manager';
      var isDirector = pos.indexOf('director') !== -1 || role === 'director' || role === 'admin';
      if (!isChef && !isManager && !isDirector) continue;

      var outletOk = isDirector || _kitchenOutletMatches_(p.outlet, outlet);
      if (!outletOk) continue;

      var key = email.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;

      out.to.push(email);
      if (isChef) out.chefs.push(email);
      if (isManager || isDirector) out.admins.push(email);
    }
  } catch (e) {
    Logger.log('_cookingAlertRecipients_ error: ' + e.toString());
  }
  return out;
}

function _sendCookAlertEmail_(outlet, items, isReminder, recipients) {
  if (!recipients || !recipients.to || !recipients.to.length) {
    Logger.log('Cook alert: no Chef/Manager/Director recipients for outlet "' + outlet + '"');
    return false;
  }
  try {
    var tz = _appTimeZone_();
    var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
    var outletLbl = outlet || 'All outlets';
    var subject = (isReminder ? '\u23F0 REMINDER: ' : '\uD83C\uDF73 ') +
      'Cook more needed \u2014 ' + outletLbl +
      ' (' + items.length + ' item' + (items.length === 1 ? '' : 's') + ')';

    var rows = items.map(function (it) {
      var out = it.stockStatus === 'out';
      var sev = out ? 'SOLD OUT' : 'Running low';
      var sevColor = out ? '#dc2626' : '#d97706';
      return '<tr>' +
        '<td style="padding:6px 10px;border:1px solid #eee;">' + _cookAlertEsc_(it.name || it.productCode) + '</td>' +
        '<td style="padding:6px 10px;border:1px solid #eee;text-align:center;color:#fff;background:' + sevColor + ';font-weight:700;">' + sev + '</td>' +
        '<td style="padding:6px 10px;border:1px solid #eee;text-align:right;">' + (Number(it.remaining) || 0) + '</td>' +
        '<td style="padding:6px 10px;border:1px solid #eee;text-align:right;">' + (Number(it.recookThreshold) || 0) + '</td>' +
        '</tr>';
    }).join('');

    var body = '<div style="font-family:Arial,sans-serif;color:#222;">' +
      '<h2 style="color:#b45309;margin:0 0 6px;">\uD83C\uDF73 Kitchen: please cook more</h2>' +
      '<p style="margin:0 0 12px;">Outlet: <b>' + _cookAlertEsc_(outletLbl) + '</b> \u00B7 ' + nowStr +
      (isReminder ? ' \u00B7 <b style="color:#dc2626;">reminder</b>' : '') + '</p>' +
      '<table style="border-collapse:collapse;border:1px solid #eee;">' +
      '<tr style="background:#fef3c7;">' +
      '<th style="padding:6px 10px;border:1px solid #eee;text-align:left;">Item</th>' +
      '<th style="padding:6px 10px;border:1px solid #eee;">Status</th>' +
      '<th style="padding:6px 10px;border:1px solid #eee;">Remaining</th>' +
      '<th style="padding:6px 10px;border:1px solid #eee;">Threshold</th></tr>' +
      rows + '</table>' +
      '<p style="margin:12px 0 0;color:#666;font-size:12px;">Reminders are sent every 30 minutes until a cooking batch is logged for these items.</p>' +
      '</div>';

    MailApp.sendEmail({ to: recipients.to.join(','), subject: subject, htmlBody: body });
    return true;
  } catch (e) {
    Logger.log('Cook alert email failed: ' + e.toString());
    return false;
  }
}

function _cookAlertEsc_(s) {
  if (s == null) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Install the recurring sweep trigger (run once by an admin after deploy). */
function installCookingStockAlertTrigger() {
  removeCookingStockAlertTrigger();
  ScriptApp.newTrigger('runCookingStockAlertSweep')
    .timeBased()
    .everyMinutes(COOK_ALERT_SWEEP_MINUTES)
    .create();
  Logger.log('Installed cooking stock alert trigger (every ' + COOK_ALERT_SWEEP_MINUTES + ' minutes)');
}

function removeCookingStockAlertTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (tg) {
    if (tg.getHandlerFunction() === 'runCookingStockAlertSweep') {
      ScriptApp.deleteTrigger(tg);
    }
  });
}
