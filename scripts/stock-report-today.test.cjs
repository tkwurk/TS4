const assert = require('node:assert/strict')
const { readFileSync } = require('node:fs')
const path = require('node:path')
const vm = require('node:vm')
const { test } = require('node:test')

const now = '2026-11-19T05:00:00Z' // 2026-11-19 12:00 in Asia/Jakarta
class TestDate extends Date {
  constructor(...args) { super(...(args.length ? args : [now])) }
  static now() { return new Date(now).getTime() }
}

function app() {
  const tables = {
    PRODUCTS: [
      ['Name', 'Category', 'Type', 'Product Code', 'BuyUnit', 'Active'],
      ['Rice', 'Food', 'Buy', 'RICE', 'kg', 'Y'],
      ['Sauce', 'Food', 'Buy', 'SAUCE', 'kg', 'Y'],
      ['Meal', 'Food', 'Sell', 'MEAL', '', 'Y'],
    ],
    STOCK_HISTORY: [
      ['AsOfDate', 'Outlet', 'ProductCode', 'TotalPurchased', 'TotalConsumed', 'TotalSold', 'UpdatedAt'],
      ['2026-11-18', 'Main', 'RICE', 10, 1, 2, ''],
    ],
    PURCHASES: [
      ['Timestamp', 'Product Code', 'Quantity', 'RQuantity', 'Outlet'],
      [new TestDate('2026-11-17T05:00:00Z'), 'RICE', 999, 999, 'Main'], // must never be read
      [new TestDate('2026-11-18T05:00:00Z'), 'RICE', 999, 999, 'Main'], // earlier date: stop marker
      [new TestDate('2026-11-19T01:00:00Z'), 'RICE', 5, 4, 'Main'],     // today: RQuantity preferred
      [new TestDate('2026-11-19T02:00:00Z'), 'RICE', 7, 7, 'West'],     // other outlet
      [new TestDate('2026-11-19T03:00:00Z'), 'SAUCE', 2, 2, 'Main'],
    ],
    ORDERS: [
      ['Timestamp', 'Order Number', 'Product Code', 'Quantity', 'Outlet'],
      [new TestDate('2026-11-18T05:00:00Z'), 'old', 'MEAL', 50, 'Main'],
      [new TestDate('2026-11-19T01:00:00Z'), 'one', 'MEAL', 2, 'Main'],
      [new TestDate('2026-11-19T02:00:00Z'), 'two', 'MEAL', 1, 'West'],
    ],
    BOM: [
      ['BOM Id', 'BOM Item Id', 'Recipe Item Code', 'Item Quantity', 'Unit Product Code', 'Status'],
      ['1', '1', 'RICE', 2, 'MEAL', 'Active'],
      ['2', '2', 'SAUCE', 0.5, 'MEAL', 'Active'],
    ],
    BOM_USAGE: [],
    INVENTORY_CONSUMPTION: [
      ['Timestamp', 'Product Code', 'Quantity Consumed', 'Outlet', 'Username'],
      [new TestDate('2026-11-18T05:00:00Z'), 'RICE', 99, 'Main', 'operator'],
      [new TestDate('2026-11-19T01:00:00Z'), 'RICE', 3, 'Main', 'operator'],
    ],
  }
  const windows = []
  const fullReads = []
  const cache = new Map()
  const session = { username: 'operator', deviceId: 'device', outlet: 'Main', allowedTabs: ['stockReport'] }
  const context = vm.createContext({
    Date: TestDate, Map,
    SHEETS: {},
    Logger: { log() {} },
    Session: { getScriptTimeZone: () => 'Asia/Jakarta' },
    getUserSessionEnhanced: (device, username) => device === 'device' && username === 'operator' ? session : null,
    CacheService: { getScriptCache: () => ({
      get: (key) => cache.get(key) ?? null,
      put: (key, value) => cache.set(key, value),
      remove: (key) => cache.delete(key),
    }) },
    Utilities: {
      getUuid: () => 'test-request',
      formatDate(date, timeZone, format) {
        assert.equal(format, 'yyyy-MM-dd')
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
          timeZone, year: 'numeric', month: '2-digit', day: '2-digit',
        }).formatToParts(date).map((part) => [part.type, part.value]))
        return `${parts.year}-${parts.month}-${parts.day}`
      },
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: (text) => ({ setMimeType: () => JSON.parse(text) }),
    },
    getGlobalSpreadsheet: () => { throw new Error('Direct spreadsheet access is forbidden') },
  })
  for (const file of [
    'utils.js', 'accessControl.js', 'repo.gs', 'products.js', 'bom.js', 'consumptionPerformance.js',
    'inventoryConsumption.js', 'stockHistory.js', 'inventoryWorker.gs', 'orchestrator.gs', 'apiRouter.gs',
  ]) {
    vm.runInContext(readFileSync(path.join(__dirname, '..', 'apps-script-source', file), 'utf8'), context, { filename: file })
  }
  context.getAppTimeZone_ = () => 'Asia/Jakarta'
  context.formatDateForComparison = (date) => context.getAppDateKey_(date)
  context.CONSUMPTION_ORDER_BATCH_SIZE = 2
  context.Repo.adapter = {
    getValues(name) {
      fullReads.push(name)
      if (name === 'PURCHASES' || name === 'ORDERS') {
        throw new Error('Full read of append-ordered sheet in live report path: ' + name)
      }
      return tables[name] ?? []
    },
    getValuesFromEnd(name, offset, count, anchorTotalRows) {
      windows.push([name, Number(offset) || 0])
      const values = tables[name] ?? []
      const headers = values[0] ?? []
      const availableRows = Math.max(0, values.length - 1)
      const totalRows = anchorTotalRows === undefined
        ? availableRows
        : Math.min(availableRows, Math.max(0, Number(anchorTotalRows) || 0))
      const endIndex = Math.max(0, totalRows - Math.max(0, Number(offset) || 0))
      const startIndex = Math.max(0, endIndex - Math.max(0, Number(count) || 0))
      return { headers, rows: values.slice(startIndex + 1, endIndex + 1), startRow: startIndex + 2, totalRows }
    },
  }
  context.WorkflowLog.record = () => {}
  context.ErrorHandler = { wrap(step, run) {
    try { return run() } catch (error) { return context.apiFail(error.message, 'TEST_ERROR') }
  } }
  function request() {
    return context.doPost({ postData: { contents: JSON.stringify({
      eventType: 'stock.report.requested',
      payload: { _authContext: { username: 'operator', deviceId: 'device', outlet: 'Main' } },
    }) } })
  }
  return { request, windows, fullReads }
}

test('report equals Stock_History snapshot plus live current-day purchases, consumption, and orders x BOM', () => {
  const instance = app()
  const result = instance.request()
  assert.equal(result.success, true, result.error)
  const step = result.data.steps.stockReport
  assert.equal(step.success, true, step.error)
  assert.equal(step.data.historyAsOfDate, '2026-11-18')
  assert.equal(step.data.todayDate, '2026-11-19')
  assert.equal(step.data.stockSource, 'History+Today')
  const byCode = Object.fromEntries(step.data.items.map((item) => [item.productCode, item]))
  // RICE: snapshot 10/1/2 + today RQuantity 4, consumption 3, orders 2xMEAL -> BOM 4
  assert.deepEqual(
    [byCode.RICE.purchased, byCode.RICE.consumed, byCode.RICE.sold, byCode.RICE.remaining],
    [14, 4, 6, 4],
  )
  // SAUCE: no snapshot + today purchase 2, orders 2xMEAL x 0.5
  assert.deepEqual(
    [byCode.SAUCE.purchased, byCode.SAUCE.consumed, byCode.SAUCE.sold, byCode.SAUCE.remaining],
    [2, 0, 1, 1],
  )
})

test('purchases and orders use bottom-up tail scans that stop before older rows', () => {
  const instance = app()
  assert.equal(instance.request().success, true)
  // Full getValues was never used for the append-ordered sheets.
  assert.ok(!instance.fullReads.includes('PURCHASES'))
  assert.ok(!instance.fullReads.includes('ORDERS'))
  // Purchases (5 rows, batch 2): offsets 0 and 2 are read; the scan stops at the
  // 2026-11-18 row, so the oldest batch (offset 4) is never requested.
  const purchaseOffsets = instance.windows.filter(([name]) => name === 'PURCHASES').map(([, offset]) => offset)
  assert.ok(purchaseOffsets.includes(0) && purchaseOffsets.includes(2), String(purchaseOffsets))
  assert.ok(!purchaseOffsets.includes(4), 'tail scan read past the first earlier business date')
  const orderOffsets = instance.windows.filter(([name]) => name === 'ORDERS').map(([, offset]) => offset)
  assert.ok(orderOffsets.includes(0), String(orderOffsets))
})
