const assert = require('node:assert/strict')
const { readFileSync } = require('node:fs')
const path = require('node:path')
const vm = require('node:vm')
const { test } = require('node:test')

function setup() {
  const properties = new Map()
  const triggers = []
  const calls = []
  const context = vm.createContext({
    SHEETS: { STOCK_HISTORY: 'Stock_History' },
    Session: { getScriptTimeZone: () => 'Asia/Jakarta' },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => properties.get(key) ?? null,
        setProperty: (key, value) => properties.set(key, String(value)),
        deleteProperty: (key) => properties.delete(key),
      }),
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (trigger) => {
        const index = triggers.indexOf(trigger)
        if (index >= 0) triggers.splice(index, 1)
      },
      newTrigger: (handler) => {
        const trigger = {
          handler,
          hour: null,
          getHandlerFunction() { return this.handler },
          getUniqueId() { return 'trigger-id' },
          getEventType() { return 'CLOCK' },
          getTriggerSource() { return 'CLOCK' },
        }
        const builder = {
          timeBased() { return this },
          atHour(hour) { trigger.hour = hour; return this },
          nearMinute() { return this },
          everyDays() { return this },
          inTimezone(timeZone) { trigger.timeZone = timeZone; return this },
          create() { triggers.push(trigger); return trigger },
        }
        return builder
      },
    },
    requirePermission_: (payload, permission) => {
      calls.push({ payload, permission })
      if (!payload || !payload._authContext) throw new Error('Authentication required')
      if (permission !== 'adminOrders') throw new Error('Admin permission required')
      return { username: 'admin' }
    },
    apiOk: (data, code) => ({ success: true, data, error: null, code }),
    apiFail: (error, code) => ({ success: false, data: null, error: String(error), code }),
    ErrorHandler: {
      wrap: (step, fn) => fn(),
    },
  })

  for (const file of [
    'apps-script-source/stockHistory.js',
    'apps-script-source/inventoryWorker.gs',
    'apps-script-source/orchestrator.gs',
  ]) {
    vm.runInContext(readFileSync(path.join(__dirname, '..', file), 'utf8'), context, { filename: file })
  }
  context.WorkflowLog.record = () => {}
  return { context, calls, triggers, properties }
}

test('scheduling replaces the trigger and exposes its hour and timezone', () => {
  const { context, triggers, properties } = setup()

  context.scheduleNightlyStockHistoryTrigger(0)
  context.scheduleNightlyStockHistoryTrigger(5)

  assert.equal(triggers.length, 1)
  assert.equal(triggers[0].hour, 5)
  assert.equal(triggers[0].timeZone, 'Asia/Jakarta')
  assert.equal(properties.get('STOCK_HISTORY_SCHEDULE_HOUR'), '5')
  const info = context.getNightlyStockHistoryTriggerInfo()
  assert.equal(info.triggerActive, true)
  assert.equal(info.scheduleHour, 5)
  assert.equal(info.timeZone, 'Asia/Jakarta')
})

test('PWA schedule event requires adminOrders and rejects invalid hours', () => {
  const { context, calls, triggers } = setup()
  const payload = { _authContext: { username: 'admin', deviceId: 'device' }, hour: 24 }

  const invalid = context.InventoryWorker.scheduleStockHistoryJob(payload, {})
  assert.equal(invalid.success, false)
  assert.equal(invalid.code, 'STOCK_HISTORY_INVALID_HOUR')
  assert.equal(triggers.length, 0)
  assert.equal(calls[0].permission, 'adminOrders')

  payload.hour = 0
  const routed = context.Orchestrator.routeEvent('stock.history.schedule.requested', payload, {
    username: 'admin',
    outlet: 'Main',
    requestId: 'request',
  })
  assert.equal(routed.success, true)
  assert.equal(routed.data.steps.stockHistorySchedule.success, true)
  assert.equal(triggers.length, 1)
  assert.equal(calls[1].permission, 'adminOrders')
})

test('PWA status event requires adminOrders and returns job status', () => {
  const { context, calls } = setup()
  context.scheduleNightlyStockHistoryTrigger(0)
  const routed = context.Orchestrator.routeEvent(
    'stock.history.status.requested',
    { _authContext: { username: 'admin', deviceId: 'device' } },
    { username: 'admin', outlet: 'Main', requestId: 'request' },
  )

  assert.equal(routed.success, true)
  assert.equal(routed.data.steps.stockHistoryStatus.data.triggerActive, true)
  assert.equal(routed.data.steps.stockHistoryStatus.data.scheduleHour, 0)
  assert.equal(calls[0].permission, 'adminOrders')
})

test('stock-history status and scheduling reject unauthenticated callers', () => {
  const { context, calls, triggers } = setup()

  assert.throws(() => context.InventoryWorker.getStockHistoryJobStatus({}, {}), /Authentication required/)
  assert.throws(() => context.InventoryWorker.scheduleStockHistoryJob({ hour: 0 }, {}), /Authentication required/)
  assert.equal(triggers.length, 0)
  assert.deepEqual(calls.map((call) => call.permission), ['adminOrders', 'adminOrders'])
})
