/**
 * repo.gs — Data Access Layer facade (Seam 2, see /AGENTS.md).
 *
 * Business logic (Orchestrator → Workers) MUST read and write data only through
 * Repo — never SpreadsheetApp and never a specific adapter directly. Repo is
 * provider-agnostic: it maps SHEETS keys to the underlying store and delegates
 * to the active adapter. By default this is SpreadsheetAppAdapter (Option C:
 * execute-as-owner, no Google Cloud, free). To switch to the service-account
 * REST path (Option B, per-user concurrency) set `Repo.adapter = SheetsAdapter`;
 * to migrate the database, implement a new adapter with the same method surface
 * (getValues / appendRows / updateValues / clearRange) and set `Repo.adapter`.
 *
 * Methods are collection-oriented and work with row OBJECTS keyed by the sheet's
 * header row, so callers don't depend on column order.
 */

var Repo = {

  /**
   * Explicit adapter override. Leave null to use the default (Option C,
   * SpreadsheetAppAdapter). Set to SheetsAdapter to switch to the service-account
   * REST path (Option B), or to any future adapter to migrate the database.
   */
  adapter: null,

  /** Resolve the active adapter lazily (avoids Apps Script file load-order issues). */
  _adapter: function () {
    if (this.adapter) return this.adapter;
    if (typeof SpreadsheetAppAdapter !== 'undefined') return SpreadsheetAppAdapter;
    if (typeof SheetsAdapter !== 'undefined') return SheetsAdapter;
    throw new Error('Repo: no data adapter available.');
  },

  /** Resolve a SHEETS key (e.g. 'ORDERS') or a literal sheet name to a name. */
  _name: function (collection) {
    if (!collection) throw new Error('Repo: collection is required.');
    return (SHEETS && SHEETS[collection]) ? SHEETS[collection] : collection;
  },

  /**
   * Read a whole collection as an array of objects keyed by the header row.
   * @param {string} collection SHEETS key or sheet name.
   * @return {{ headers: string[], rows: Object[] }}
   */
  readTable: function (collection) {
    var values = this._adapter().getValues(this._name(collection));
    if (!values.length) return { headers: [], rows: [] };
    var headers = values[0].map(function (h) { return (h == null ? '' : String(h)).trim(); });
    var rows = [];
    for (var r = 1; r < values.length; r++) {
      rows.push(this._rowToObject_(headers, values[r]));
    }
    return { headers: headers, rows: rows };
  },

  /** Convenience: just the row objects for a collection. */
  readRows: function (collection) {
    return this.readTable(collection).rows;
  },

  /**
   * Append one or more objects to a collection, aligning each object's values to
   * the existing header row.
   * @param {string} collection
   * @param {Object|Object[]} objects
   */
  appendObjects: function (collection, objects) {
    var list = Array.isArray(objects) ? objects : [objects];
    if (!list.length) return { appended: 0 };
    var headers = this.readTable(collection).headers;
    if (!headers.length) {
      throw new Error('Repo.appendObjects: collection "' + collection + '" has no header row.');
    }
    var rows = list.map(function (obj) {
      return headers.map(function (h) {
        var v = obj[h];
        return v == null ? '' : v;
      });
    });
    this._adapter().appendRows(this._name(collection), rows);
    return { appended: rows.length };
  },

  /** Append raw 2D rows when you already have ordered arrays. */
  appendRows: function (collection, rows) {
    this._adapter().appendRows(this._name(collection), rows);
    return { appended: (rows || []).length };
  },

  /** Overwrite an explicit A1 range within a collection, e.g. 'A2:F2'. */
  updateRange: function (collection, a1, rows) {
    this._adapter().updateValues(this._name(collection) + '!' + a1, rows);
  },

  _rowToObject_: function (headers, row) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = (c < row.length) ? row[c] : '';
    }
    return obj;
  }
};

/**
 * repoSelfTest — run from the Apps Script editor (Run ▸ repoSelfTest) to verify
 * the data layer reads through the active adapter. Read-only; writes nothing.
 * Reads the Products sheet and logs the active adapter, header count, and row
 * count. Returns a short status string.
 */
function repoSelfTest() {
  try {
    var adapterName = (Repo._adapter() === (typeof SpreadsheetAppAdapter !== 'undefined' ? SpreadsheetAppAdapter : null))
      ? 'SpreadsheetAppAdapter (Option C)'
      : 'SheetsAdapter (Option B)';
    var table = Repo.readTable('PRODUCTS');
    var msg = 'Repo OK via ' + adapterName + ' — headers: ' + table.headers.length +
      ', rows: ' + table.rows.length;
    Logger.log(msg);
    return msg;
  } catch (e) {
    var err = 'repoSelfTest FAILED: ' + e.toString();
    Logger.log(err);
    return err;
  }
}
