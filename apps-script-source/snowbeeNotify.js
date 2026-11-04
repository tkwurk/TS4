/**
 * Snowbee UAT notification + CSV exports (E1.5 / G2).
 */

function sendSnowbeeUATReadyEmail(recipientEmail) {
  var props = PropertiesService.getScriptProperties();
  var to = recipientEmail || props.getProperty('SNOWBEE_TEST_EMAIL') || 'tehsinuk@gmail.com';

  var stagingUrl = props.getProperty('SNOWBEE_STAGING_URL') ||
    'https://script.google.com/macros/s/AKfycbzOiTKLAEOoVHN2XTIoXyNa7n03LTNlmuABMrCjT8TV/exec';

  var body = [
    'Hi — POS_TS3 Snowbee UAT is ready for your review.',
    '',
    'What is included (local repo, deploy to staging when ready):',
    '• A3.3–A3.5: header validation, apiOk/apiFail, atomic order save',
    '• E1: P&L bundle + cache (not E2 archive)',
    '• B1 foundation: People sheet, roles.js, accessControl.js, login returns role',
    '• B3: AuditLog sheet + writeAuditLog on person save',
    '• C4: getBusinessDate() business-day helper',
    '• D1 partial: purchasingReceive/checklist keys, label fixes, role badge',
    '• F1: Admin → People tab (list view)',
    '• F2.1: inactive products hidden from POS',
    '',
    'Staging URL (@HEAD): ' + stagingUrl,
    '',
    'Before testing:',
    '1. clasp push --project .clasp.staging.json (when approved)',
    '2. In Apps Script editor run once: migrateUsersAndEmployees()',
    '3. Run smokeTestSnowbee() in editor',
    '',
    'Smoke checklist:',
    '• Login (cashier / supervisor / admin)',
    '• Checkout cash order',
    '• Language ID ↔ EN on sidebar',
    '• P&L tab switches after first load',
    '• Admin → People tab',
    '• Deposit + expense',
    '',
    '— Snowbee / POS_TS3'
  ].join('\n');

  MailApp.sendEmail({
    to: to,
    subject: '[POS_TS3 Snowbee] UAT ready — please test',
    body: body
  });

  return apiOk({ sentTo: to }, 'UAT_EMAIL_SENT');
}

/** Run from Apps Script editor or `clasp run sendSnowbeeUATReadyEmailNow` */
function sendSnowbeeUATReadyEmailNow() {
  return sendSnowbeeUATReadyEmail('tehsinuk@gmail.com');
}

function exportSalesSummaryCsvSecure(params) {
  try {
    var session = requireAuth_(params);
    var dateStr = params.date || getBusinessDateKey(new Date());
    var summaryFn = typeof getSalesSummaryByDate === 'function' ? getSalesSummaryByDate : null;
    if (!summaryFn) return apiFail('Sales summary not available', 'EXPORT_UNAVAILABLE');
    var raw = summaryFn(dateStr, session.username);
    var parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    var rows = [['Product', 'Qty', 'Total']];
    (parsed.data || []).forEach(function(r) {
      rows.push([r.product || r.name || '', r.qty || 0, r.total || 0]);
    });
    var csv = rows.map(function(r) {
      return r.map(function(c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(',');
    }).join('\n');
    return apiOk({ csv: csv, filename: 'sales-' + dateStr + '.csv' }, 'CSV_EXPORT');
  } catch (e) {
    return apiFail(e.message || e, 'CSV_EXPORT_FAILED');
  }
}

function exportPnLCsvSecure(params) {
  try {
    var session = requireAuth_(params);
    var year = Number(params.year) || new Date().getFullYear();
    var bundle = getPnLBundle(session.deviceId, session.username, year);
    var rows = [['Month', 'Sales', 'Purchases', 'Expenses', 'Profit']];
    (bundle.monthly || []).forEach(function(m) {
      rows.push([m.label, m.sales, m.purchases, m.expenses, m.profit]);
    });
    var csv = rows.map(function(r) {
      return r.map(function(c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(',');
    }).join('\n');
    return apiOk({ csv: csv, filename: 'pnl-' + year + '.csv' }, 'CSV_EXPORT');
  } catch (e) {
    return apiFail(e.message || e, 'CSV_EXPORT_FAILED');
  }
}

function getTriggerHealthSecure(params) {
  try {
    requireAdmin_(params);
    var triggers = ScriptApp.getProjectTriggers();
    var list = triggers.map(function(t) {
      return {
        handler: t.getHandlerFunction(),
        type: String(t.getEventType()),
        source: t.getTriggerSourceId ? t.getTriggerSourceId() : ''
      };
    });
    return apiOk({ triggers: list, count: list.length }, 'TRIGGER_HEALTH');
  } catch (e) {
    return apiFail(e.message || e, 'TRIGGER_HEALTH_FAILED');
  }
}
