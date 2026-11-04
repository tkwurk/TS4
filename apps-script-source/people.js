/**
 * People sheet — unified login + HR (Snowbee B1 / F1).
 */

var PEOPLE_HEADER_ROW = [
  'person_id', 'display_name', 'pin_code', 'outlet', 'role', 'email',
  'department', 'position', 'employment_type', 'active', 'can_login',
  'extra_tabs', 'denied_tabs'
];

function ensurePeopleSheet_() {
  var ss = getGlobalSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PEOPLE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PEOPLE);
    sheet.appendRow(PEOPLE_HEADER_ROW);
    Logger.log('Created People sheet');
  }
  return sheet;
}

function personFromRow_(row) {
  if (!row || !row[0]) return null;
  return {
    person_id: String(row[0]).trim(),
    display_name: row[1] || '',
    pin_code: row[2] != null ? String(row[2]) : '',
    outlet: row[3] || '',
    role: normalizeRole_(row[4] || 'cashier'),
    email: row[5] || '',
    department: row[6] || '',
    position: row[7] || '',
    employment_type: row[8] || '',
    active: (row[9] || 'Y').toString().toUpperCase(),
    can_login: (row[10] || 'Y').toString().toUpperCase(),
    extra_tabs: row[11] || '',
    denied_tabs: row[12] || ''
  };
}

function rowFromPerson_(person) {
  return [
    person.person_id,
    person.display_name || person.person_id,
    person.pin_code || '',
    person.outlet || '',
    normalizeRole_(person.role || 'cashier'),
    person.email || '',
    person.department || '',
    person.position || '',
    person.employment_type || '',
    (person.active || 'Y').toString().toUpperCase(),
    (person.can_login || 'Y').toString().toUpperCase(),
    person.extra_tabs || '',
    person.denied_tabs || ''
  ];
}

function getPersonById_(personId) {
  ensurePeopleSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PEOPLE);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(personId).trim()) {
      return personFromRow_(data[i]);
    }
  }
  return null;
}

function getPersonForLogin_(personId, pin) {
  var person = getPersonById_(personId);
  if (!person) return null;
  if (person.active === 'N') return null;
  if (person.can_login === 'N') return null;
  if (String(person.pin_code) !== String(pin)) return null;
  return person;
}

/** Live auth profile from People (no PIN). */
function getAuthProfileForUsername_(username) {
  var person = getPersonById_(username);
  if (!person || person.active === 'N') return null;
  return {
    username: person.person_id,
    outlet: person.outlet,
    role: normalizeRole_(person.role),
    allowedTabs: normalizeAllowedTabsForRole_(resolveEffectiveTabs(person), person.role),
    can_login: person.can_login,
    displayName: person.display_name
  };
}

function refreshSessionAuthFromPeople_(username, session) {
  var profile = getAuthProfileForUsername_(username);
  if (!session) return session;
  if (profile) {
    session.role = profile.role;
    session.allowedTabs = normalizeAllowedTabsForRole_(profile.allowedTabs, profile.role);
    if (profile.outlet) session.outlet = profile.outlet;
  }
  session.outlet = resolveUserOutlet_(username, session.outlet);
  return session;
}

/** People → provided value; empty string if none found. */
function resolveUserOutlet_(username, preferredOutlet) {
  var o = (preferredOutlet || '').toString().trim();
  if (o && o !== 'undefined') return o;
  var profile = getAuthProfileForUsername_(username);
  if (profile && profile.outlet) {
    o = profile.outlet.toString().trim();
    if (o) return o;
  }
  return '';
}

function inferRoleFromAllowedTabs_(allowedTabsStr) {
  var tabs = parseCsvList_(allowedTabsStr);
  if (tabs.indexOf('adminOrders') >= 0 || tabs.indexOf('*') >= 0) return 'admin';
  if (tabs.indexOf('purchasing') >= 0 && tabs.indexOf('profitLoss') >= 0) return 'supervisor';
  if (tabs.indexOf('investments') >= 0) return 'manager';
  if (tabs.indexOf('orders') >= 0) return 'cashier';
  return 'readonly';
}

/**
 * One-time migration: Users + employees → People.
 * Safe to run multiple times (skips existing person_id).
 */
function migrateUsersAndEmployees() {
  ensurePeopleSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PEOPLE);
  var existing = {};
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) existing[String(data[i][0]).trim()] = true;
  }

  var added = 0;
  var usersSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.USERS);
  if (usersSheet) {
    var users = usersSheet.getDataRange().getValues();
    for (var u = 1; u < users.length; u++) {
      var username = getFixedRowField(users[u], 'USERS', 'USERNAME');
      if (!username || existing[username]) continue;
      var allowedTabs = getFixedRowField(users[u], 'USERS', 'ALLOWED_TABS') || '';
      var person = {
        person_id: username,
        display_name: username,
        pin_code: getFixedRowField(users[u], 'USERS', 'CODE'),
        outlet: getFixedRowField(users[u], 'USERS', 'OUTLET') || '',
        role: inferRoleFromAllowedTabs_(allowedTabs),
        email: '',
        active: 'Y',
        can_login: 'Y',
        extra_tabs: allowedTabs ? '' : '',
        denied_tabs: ''
      };
      if (allowedTabs && !ROLE_DEFAULT_TABS[person.role]) {
        person.extra_tabs = allowedTabs;
      }
      sheet.appendRow(rowFromPerson_(person));
      existing[username] = true;
      added++;
    }
  }

  var empSheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EMPLOYEES);
  if (empSheet) {
    var emps = empSheet.getDataRange().getValues();
    for (var e = 1; e < emps.length; e++) {
      var empId = emps[e][0];
      if (!empId || existing[String(empId).trim()]) continue;
      sheet.appendRow(rowFromPerson_({
        person_id: String(empId).trim(),
        display_name: emps[e][1] || empId,
        pin_code: '',
        outlet: '',
        role: 'readonly',
        email: emps[e][7] || '',
        department: emps[e][2] || '',
        position: emps[e][3] || '',
        employment_type: emps[e][4] || '',
        active: 'Y',
        can_login: 'N'
      }));
      existing[String(empId).trim()] = true;
      added++;
    }
  }

  return apiOk({ added: added, message: 'Migration complete' }, 'PEOPLE_MIGRATED');
}

function listPeopleForAdmin_(session) {
  ensurePeopleSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PEOPLE);
  var data = sheet.getDataRange().getValues();
  var people = [];
  for (var i = 1; i < data.length; i++) {
    var p = personFromRow_(data[i]);
    if (p) {
      people.push({
        person_id: p.person_id,
        display_name: p.display_name,
        outlet: p.outlet,
        role: p.role,
        email: p.email,
        department: p.department,
        position: p.position,
        active: p.active,
        can_login: p.can_login,
        extra_tabs: p.extra_tabs,
        denied_tabs: p.denied_tabs
      });
    }
  }
  return people;
}

function getPeopleForAdminSecure(params) {
  try {
    var session = requireAdmin_(params);
    return apiOk({ people: listPeopleForAdmin_(session) }, 'PEOPLE_LIST');
  } catch (e) {
    return apiFail(e.message || e, 'PEOPLE_LIST_FAILED');
  }
}

function savePersonSecure(params) {
  try {
    var session = requireAdmin_(params);
    var person = params.person || params;
    if (!person.person_id) return apiFail('person_id is required', 'PERSON_INVALID');

    ensurePeopleSheet_();
    var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PEOPLE);
    var data = sheet.getDataRange().getValues();
    var rowNum = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(person.person_id).trim()) {
        rowNum = i + 1;
        break;
      }
    }

    var row = rowFromPerson_(person);
    if (rowNum > 0) {
      var old = personFromRow_(data[rowNum - 1]);
      if (!person.pin_code && old) row[2] = old.pin_code;
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      writeAuditLog_(session.username, 'UPDATE_PERSON', person.person_id, JSON.stringify(person));
    } else {
      sheet.appendRow(row);
      writeAuditLog_(session.username, 'CREATE_PERSON', person.person_id, JSON.stringify(person));
    }

    return apiOk({ person_id: person.person_id }, 'PERSON_SAVED');
  } catch (e) {
    return apiFail(e.message || e, 'PERSON_SAVE_FAILED');
  }
}

function validateLoginPersonOrUser_(username, code) {
  var person = getPersonForLogin_(username, code);
  if (person) {
    return {
      success: true,
      username: person.person_id,
      outlet: resolveUserOutlet_(person.person_id, person.outlet),
      role: person.role,
      allowedTabs: resolveEffectiveTabs(person),
      displayName: person.display_name
    };
  }
  var existing = getPersonById_(username);
  if (existing && existing.can_login === 'N') {
    return { success: false, message: 'Login disabled for this account' };
  }
  return { success: false, message: 'Invalid credentials' };
}

function getEmployeeListFromPeople(deviceId, username) {
  var session = getUserSessionEnhanced(deviceId, username) || {};
  ensurePeopleSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PEOPLE);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var p = personFromRow_(data[i]);
    if (!p || p.active === 'N') continue;
    if (session.outlet && p.outlet && p.outlet !== session.outlet) continue;
    list.push({
      employee_id: p.person_id,
      name: p.display_name,
      department: p.department,
      position: p.position,
      employment_type: p.employment_type,
      email: p.email
    });
  }
  if (list.length) return list;
  return getEmployeeListFromEmployeesSheet_();
}

function getEmployeeListFromEmployeesSheet_() {
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var cols = buildColumnMap(headers, 'EMPLOYEES');
  var employees = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empId = row[cols.EMPLOYEE_ID >= 0 ? cols.EMPLOYEE_ID : 0];
    var empflag = getFixedRowField(row, 'EMPLOYEES', 'INACTIVE_FLAG');
    if (!empId || empflag === 'X') continue;
    employees.push({
      employee_id: empId.toString(),
      name: row[cols.NAME >= 0 ? cols.NAME : 1] || '',
      department: row[cols.DEPARTMENT >= 0 ? cols.DEPARTMENT : 2] || '',
      position: row[cols.POSITION >= 0 ? cols.POSITION : 3] || ''
    });
  }
  return employees;
}
