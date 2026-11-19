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

  getValuesFromEnd: function (name, offset, count, anchorTotalRows) {
    var values = this.getValues(name);
    if (!values.length) {
      return { headers: [], rows: [], startRow: 2, totalRows: 0 };
    }

    var headers = values[0].map(function (h) { return (h == null ? '' : String(h)).trim(); });
    var availableRows = Math.max(0, values.length - 1);
    var totalRows = anchorTotalRows === undefined
      ? availableRows
      : Math.min(availableRows, Math.max(0, Number(anchorTotalRows) || 0));
    var endIndex = Math.max(0, totalRows - Math.max(0, Number(offset) || 0));
    var startIndex = Math.max(0, endIndex - Math.max(0, Number(count) || 0));
    return {
      headers: headers,
      rows: values.slice(startIndex + 1, endIndex + 1),
      startRow: startIndex + 2,
      totalRows: totalRows
    };
  },

  ensureTable: function (name, headers) {
    var base = this._base();
    var metadata = this._fetch(base + '?fields=' + encodeURIComponent('sheets.properties.title'), { method: 'get' });
    var exists = (metadata.sheets || []).some(function (sheet) {
      return sheet.properties && sheet.properties.title === name;
    });
    if (!exists) {
      this._fetch(base + ':batchUpdate', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ requests: [{ addSheet: { properties: { title: name } } }] })
      });
    }

    var current = this.getValues(name);
    if (!current.length) {
      var endColumn = headers.length;
      var columnName = '';
      while (endColumn > 0) {
        var remainder = (endColumn - 1) % 26;
        columnName = String.fromCharCode(65 + remainder) + columnName;
        endColumn = Math.floor((endColumn - 1) / 26);
      }
      this.updateValues(name + '!A1:' + columnName + '1', [headers]);
      return true;
    }

    for (var i = 0; i < headers.length; i++) {
      if (String(current[0][i] || '').trim() !== String(headers[i] || '').trim()) {
        throw new Error('SheetsAdapter.ensureTable: header mismatch in ' + name);
      }
    }
    return true;
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
