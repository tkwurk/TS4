/**
 * Business day cutoff (Snowbee C4).
 * Default: 08:00 Asia/Jakarta — transactions before cutoff belong to previous business day.
 */
var BUSINESS_DAY_CUTOFF_HOUR = 8;

function getBusinessDate(referenceDate) {
  var ref = referenceDate ? new Date(referenceDate) : new Date();
  var tz = Session.getScriptTimeZone();
  var local = new Date(ref.toLocaleString('en-US', { timeZone: tz }));
  if (local.getHours() < BUSINESS_DAY_CUTOFF_HOUR) {
    local.setDate(local.getDate() - 1);
  }
  local.setHours(0, 0, 0, 0);
  return local;
}

function getBusinessDateKey(referenceDate) {
  var d = getBusinessDate(referenceDate);
  return d.getFullYear() + '-' +
    ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
    ('0' + d.getDate()).slice(-2);
}

function isSameBusinessDay_(dateA, dateB) {
  return getBusinessDateKey(dateA) === getBusinessDateKey(dateB);
}

function getSupervisorSnapshotSecure(params) {
  try {
    var session = requireAuth_(params);
    var todayKey = getBusinessDateKey(new Date());
    var bundle = getPnLBundle(session.deviceId, session.username, new Date().getFullYear());
    var daily = (bundle.daily || []).filter(function(d) { return d.date === todayKey; });
    var sales = 0;
    daily.forEach(function(d) { sales += d.sales || 0; });
    var cashInHand = typeof getTotalCashInHand === 'function'
      ? getTotalCashInHand(session.deviceId, session.username)
      : null;
    return apiOk({
      businessDate: todayKey,
      todaySales: sales,
      cashInHand: cashInHand,
      outlet: session.outlet
    }, 'SUPERVISOR_SNAPSHOT');
  } catch (e) {
    return apiFail(e.message || e, 'SNAPSHOT_FAILED');
  }
}
