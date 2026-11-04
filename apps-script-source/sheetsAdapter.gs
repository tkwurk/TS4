/**
 * sheetsAdapter.gs — Google Sheets implementation of the data-access adapter.
 *
 * ARCHITECTURE (see /AGENTS.md): this is the ONLY place that knows the data lives
 * in Google Sheets. It uses the Sheets REST API authenticated with the
 * ServiceAccount token (NOT SpreadsheetApp), so the backing spreadsheet stays
 * hidden from end users while the Web App runs "as the user accessing".
 *
 * To migrate to another database later, write a new adapter exposing the same
 * methods (getValues / appendRows / updateValues / clearRange) and point
 * Repo.adapter at it. Business logic never changes.
 *
 * `range` arguments are A1 notation, e.g. "Orders" (whole sheet) or "Orders!A2:F".
 */

var SheetsAdapter = {

  _base: function () {
    if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID is not set.');
    return 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(SPREADSHEET_ID);
  },

  _fetch: function (url, options) {
    options = options || {};
    options.muteHttpExceptions = true;
    options.headers = options.headers || {};
    options.headers.Authorization = 'Bearer ' + ServiceAccount.getAccessToken();

    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error('SheetsAdapter ' + (options.method || 'get') + ' failed (' + code + ') for ' + url + ': ' + body);
    }
    return body ? JSON.parse(body) : {};
  },

  /** @return {Array<Array>} 2D values for the range (empty array if none). */
  getValues: function (range) {
    var url = this._base() + '/values/' + encodeURIComponent(range);
    var out = this._fetch(url, { method: 'get' });
    return out.values || [];
  },

  /** Append rows to the table anchored at `range`. */
  appendRows: function (range, rows) {
    if (!rows || !rows.length) return { updates: { updatedRows: 0 } };
    var url = this._base() + '/values/' + encodeURIComponent(range) +
      ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS';
    return this._fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ values: rows })
    });
  },

  /** Overwrite the exact `range` with `rows`. */
  updateValues: function (range, rows) {
    var url = this._base() + '/values/' + encodeURIComponent(range) + '?valueInputOption=USER_ENTERED';
    return this._fetch(url, {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify({ values: rows })
    });
  },

  /** Clear all values in `range`. */
  clearRange: function (range) {
    var url = this._base() + '/values/' + encodeURIComponent(range) + ':clear';
    return this._fetch(url, { method: 'post' });
  }
};
