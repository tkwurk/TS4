/**
 * spreadsheetAppAdapter.gs — SpreadsheetApp implementation of the data adapter.
 *
 * ARCHITECTURE (see /AGENTS.md): this is the DEFAULT data adapter (Option C).
 * The Web App deploys "execute as the OWNER (me)", so the script reaches the
 * spreadsheet with the owner's access — end users never need (and don't get)
 * direct access to it. No Google Cloud project or service account required; the
 * app stays entirely free.
 *
 * Exposes the same method surface as SheetsAdapter (getValues / appendRows /
 * updateValues / clearRange) so the two are interchangeable behind Repo. To move
 * to per-user concurrency later, set `Repo.adapter = SheetsAdapter` (Option B).
 *
 * `range` is A1 notation: a sheet name alone ("Orders") means the whole used
 * range; "Orders!A2:F2" targets a sub-range.
 */

var SpreadsheetAppAdapter = {

  _sheet: function (name) {
    return getGlobalSpreadsheet().getSheetByName(name);
  },

  /** Split "Sheet!A1:B2" into { sheet, a1 }; a sheet name alone → a1 = null. */
  _parse: function (range) {
    var str = String(range);
    var bang = str.indexOf('!');
    if (bang === -1) return { sheet: this._unquote(str), a1: null };
    return { sheet: this._unquote(str.substring(0, bang)), a1: str.substring(bang + 1) };
  },

  _unquote: function (name) {
    name = String(name).trim();
    if (name.length >= 2 && name.charAt(0) === "'" && name.charAt(name.length - 1) === "'") {
      name = name.substring(1, name.length - 1).replace(/''/g, "'");
    }
    return name;
  },

  /** @return {Array<Array>} values; [] if the sheet is missing or empty. */
  getValues: function (range) {
    var p = this._parse(range);
    var sheet = this._sheet(p.sheet);
    if (!sheet) return [];
    if (p.a1) return sheet.getRange(p.a1).getValues();
    if (sheet.getLastRow() === 0) return [];
    return sheet.getDataRange().getValues();
  },

  appendRows: function (range, rows) {
    if (!rows || !rows.length) return { updatedRows: 0 };
    var p = this._parse(range);
    var sheet = this._sheet(p.sheet);
    if (!sheet) throw new Error('SpreadsheetAppAdapter.appendRows: sheet not found: ' + p.sheet);
    if (typeof appendRowsAtomic_ === 'function') {
      appendRowsAtomic_(sheet, rows);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { updatedRows: rows.length };
  },

  updateValues: function (range, rows) {
    var p = this._parse(range);
    var sheet = this._sheet(p.sheet);
    if (!sheet) throw new Error('SpreadsheetAppAdapter.updateValues: sheet not found: ' + p.sheet);
    if (!p.a1) throw new Error('SpreadsheetAppAdapter.updateValues requires an A1 range: ' + range);
    sheet.getRange(p.a1).setValues(rows);
    return { updatedRows: rows.length };
  },

  clearRange: function (range) {
    var p = this._parse(range);
    var sheet = this._sheet(p.sheet);
    if (!sheet) return {};
    if (p.a1) sheet.getRange(p.a1).clearContent();
    else sheet.clearContents();
    return {};
  }
};
