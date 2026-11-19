const assert = require('node:assert/strict')
const { readFileSync } = require('node:fs')
const path = require('node:path')
const vm = require('node:vm')
const { test } = require('node:test')

function loadGetDistinctOutlets(tables) {
  const context = vm.createContext({
    Repo: {
      readRows: (collection) => tables[collection] || [],
    },
  })
  vm.runInContext(
    readFileSync(path.join(__dirname, '..', 'apps-script-source', 'bom.js'), 'utf8'),
    context,
    { filename: 'apps-script-source/bom.js' },
  )
  return context.getDistinctOutletsForConsumption_()
}

test('outlet discovery uses outlet names instead of IDs', () => {
  const outlets = loadGetDistinctOutlets({
    OUTLETS: [
      { 'Outlet ID': 'outlet-17', 'Outlet Name': 'Central Kitchen' },
    ],
  })

  assert.deepEqual(Array.from(outlets), ['Central Kitchen'])
})

test('outlet discovery still includes outlets present only in transactional rows', () => {
  const outlets = loadGetDistinctOutlets({
    OUTLETS: [{ 'Outlet ID': 'outlet-17', 'Outlet Name': 'Central Kitchen' }],
    ORDERS: [{ Outlet: 'West Kitchen' }],
    PURCHASES: [{ Outlet: 'Central Kitchen' }],
    INVENTORY_CONSUMPTION: [],
  })

  assert.deepEqual(Array.from(outlets), ['Central Kitchen', 'West Kitchen'])
})
