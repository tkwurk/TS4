/**
 * TabAccess sheet — role → tab mapping (Snowbee B1).
 * Edit TabAccess sheet to change defaults without code deploy.
 *
 * Sheet columns: Role | Tabs | Description
 * - Role: lowercase role id (cashier, supervisor, manager, admin, chef, readonly)
 * - Tabs: comma-separated tab ids, or * for admin (all tabs)
 * - Optional row Role=_all_tabs lists every tab id (used when admin has * but denied_tabs)
 */

var TAB_ACCESS_CACHE_KEY = 'tab_access_map_v1';
var TAB_ACCESS_CACHE_TTL = 300;

function ensureTabAccessSheet_() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TAB_ACCESS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.TAB_ACCESS);
    sheet.getRange(1, 1, 1, 3).setValues([['Role', 'Tabs', 'Description']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    seedTabAccessFromDefaults_();
    Logger.log('Created TabAccess sheet and seeded default role tabs');
  }
  return sheet;
}

function clearTabAccessCache_() {
  try {
    CacheService.getScriptCache().remove(TAB_ACCESS_CACHE_KEY);
  } catch (e) {
    Logger.log('clearTabAccessCache_ warning: ' + e);
  }
}

function loadTabAccessMap_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(TAB_ACCESS_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (parseErr) {
      clearTabAccessCache_();
    }
  }

  var map = { roles: {}, allTabIds: [] };
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.TAB_ACCESS);
  if (!sheet || sheet.getLastRow() < 2) {
    map = buildFallbackTabAccessMap_();
    cache.put(TAB_ACCESS_CACHE_KEY, JSON.stringify(map), TAB_ACCESS_CACHE_TTL);
    return map;
  }

  var data = sheet.getDataRange().getValues();
  var tabSet = {};

  for (var i = 1; i < data.length; i++) {
    var role = (data[i][0] || '').toString().trim().toLowerCase();
    var tabsStr = (data[i][1] || '').toString().trim();
    if (!role) continue;

    if (role === '_all_tabs' || role === 'all_tabs') {
      parseCsvList_(tabsStr).forEach(function(tab) {
        if (tab && tab !== '*') tabSet[tab] = true;
      });
      continue;
    }

    if (role.charAt(0) === '_') continue;

    map.roles[role] = parseCsvList_(tabsStr);
    map.roles[role].forEach(function(tab) {
      if (tab && tab !== '*') tabSet[tab] = true;
    });
  }

  map.allTabIds = Object.keys(tabSet).sort();
  if (!map.allTabIds.length) {
    map.allTabIds = ALL_TAB_IDS.slice();
  }

  cache.put(TAB_ACCESS_CACHE_KEY, JSON.stringify(map), TAB_ACCESS_CACHE_TTL);
  return map;
}

function buildFallbackTabAccessMap_() {
  var roles = {};
  Object.keys(ROLE_DEFAULT_TABS).forEach(function(role) {
    roles[role] = ROLE_DEFAULT_TABS[role].slice();
  });
  return { roles: roles, allTabIds: ALL_TAB_IDS.slice() };
}

function getRoleTabsFromSheet_(role) {
  var roleKey = (role || 'cashier').toString().trim().toLowerCase();
  if (roleKey === 'kitchen') roleKey = 'chef';

  var map = loadTabAccessMap_();
  if (map.roles[roleKey] && map.roles[roleKey].length) {
    return map.roles[roleKey].slice();
  }
  return (ROLE_DEFAULT_TABS[roleKey] || ROLE_DEFAULT_TABS.cashier).slice();
}

function getKnownRolesFromSheet_() {
  var map = loadTabAccessMap_();
  var roles = Object.keys(map.roles || {});
  if (roles.length) return roles;
  return Object.keys(ROLE_DEFAULT_TABS);
}

function getAllTabIds_() {
  var map = loadTabAccessMap_();
  if (map.allTabIds && map.allTabIds.length) {
    return map.allTabIds.slice();
  }
  return ALL_TAB_IDS.slice();
}

/**
 * One-time / repair: write code defaults into TabAccess when rows are missing.
 * Safe to re-run — only appends roles not already present; refreshes _all_tabs if missing.
 */
function seedTabAccessFromDefaults() {
  ensureTabAccessSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.TAB_ACCESS);
  var data = sheet.getDataRange().getValues();
  var existing = {};
  var hasAllTabs = false;

  for (var i = 1; i < data.length; i++) {
    var role = (data[i][0] || '').toString().trim().toLowerCase();
    if (!role) continue;
    if (role === '_all_tabs' || role === 'all_tabs') hasAllTabs = true;
    else if (role.charAt(0) !== '_') existing[role] = true;
  }

  var added = 0;
  Object.keys(ROLE_DEFAULT_TABS).forEach(function(role) {
    if (existing[role]) return;
    sheet.appendRow([
      role,
      ROLE_DEFAULT_TABS[role].join(','),
      'Seeded from code defaults'
    ]);
    added++;
  });

  if (!hasAllTabs) {
    sheet.appendRow([
      '_all_tabs',
      ALL_TAB_IDS.join(','),
      'Catalog of all tab ids (for admin denied_tabs)'
    ]);
    added++;
  }

  clearTabAccessCache_();
  return apiOk({ added: added }, 'TAB_ACCESS_SEEDED');
}

function seedTabAccessFromDefaults_() {
  seedTabAccessFromDefaults();
}

function getTabAccessForAdminSecure(params) {
  try {
    requireAdmin_(params);
    ensureTabAccessSheet_();
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.TAB_ACCESS);
    var data = sheet.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({
        role: String(data[i][0]),
        tabs: String(data[i][1] || ''),
        description: String(data[i][2] || '')
      });
    }
    clearTabAccessCache_();
    var map = loadTabAccessMap_();
    return apiOk({
      rows: rows,
      resolved: map.roles,
      allTabIds: map.allTabIds
    }, 'TAB_ACCESS_OK');
  } catch (e) {
    return apiFail(e.message || String(e), 'TAB_ACCESS_FAILED');
  }
}
