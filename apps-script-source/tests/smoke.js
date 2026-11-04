/**
 * Smoke tests — run in Apps Script editor (Snowbee G1.1).
 */
function smokeTestSnowbee() {
  var results = [];
  function ok(name, pass, detail) {
    results.push({ name: name, pass: !!pass, detail: detail || '' });
  }

  ok('config SCHEMA_VERSION', typeof SCHEMA_VERSION === 'string', SCHEMA_VERSION);
  ok('roles.js loaded', typeof resolveEffectiveTabs === 'function');
  ok('apiOk', apiOk({ x: 1 }).success === true);
  ok('getBusinessDate', typeof getBusinessDateKey(new Date()) === 'string');
  ok('initializeSpreadsheet', initializeSpreadsheet() === true);

  var tabs = resolveEffectiveTabs({ role: 'cashier' });
  ok('cashier tabs', tabs.indexOf('orders') >= 0, tabs.join(','));
  ok('tab access loader', typeof getRoleTabsFromSheet_ === 'function');
  ok('tab access sheet key', typeof SHEETS.TAB_ACCESS === 'string', SHEETS.TAB_ACCESS);

  Logger.log(JSON.stringify(results, null, 2));
  return apiOk({ results: results }, 'SMOKE_TEST');
}
