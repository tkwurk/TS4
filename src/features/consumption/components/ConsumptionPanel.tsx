import { useCallback, useEffect, useMemo, useState } from 'react'
import type { AuthContext } from '@/services/api'
import { addDays, fmtDateOnly, localDate } from '@/lib/dates'
import { Badge, Button, DataTable, Field, FormRow, Notice, Panel, Select, Stats, SubTabs, TextInput } from '@/components/ui'
import { useI18n } from '@/i18n'
import { getConsumptionHistory, getDailyClosingBundle, recordConsumption, type ConsumptionHistoryRow, type DailyClosingBundle } from '../api'

type Tab = 'entry' | 'summary' | 'history'
type Msg = { tone: 'info' | 'error'; text: string }

/**
 * Consumption tab — mirrors ui_inventoryConsumption.html: daily closing bundle
 * (buy products + auto-prefill from today's sales + today's summary), operator
 * enters/adjusts consumed quantities per product and saves.
 */
export function ConsumptionPanel({ auth }: { auth: AuthContext }) {
  const { t, fmtNum } = useI18n()
  const [tab, setTab] = useState<Tab>('entry')
  const [bundle, setBundle] = useState<DailyClosingBundle | null>(null)
  const [loading, setLoading] = useState(false)
  const [saving, setSaving] = useState(false)
  const [qty, setQty] = useState<Record<string, string>>({})
  const [search, setSearch] = useState('')
  const [category, setCategory] = useState('')
  const [msg, setMsg] = useState<Msg | null>(null)
  const [historyStart, setHistoryStart] = useState(() => addDays(localDate(), -29))
  const [historyEnd, setHistoryEnd] = useState(() => localDate())
  const [historyRows, setHistoryRows] = useState<ConsumptionHistoryRow[]>([])
  const [historyLoading, setHistoryLoading] = useState(false)
  const [historyError, setHistoryError] = useState<string | null>(null)

  const load = useCallback(async () => {
    setLoading(true)
    const r = await getDailyClosingBundle(auth)
    setLoading(false)
    if (!r.success || !r.data) return setMsg({ tone: 'error', text: r.error ?? t('cons.loadFailed') })
    const nextData = r.data
    setCategory((current) =>
      current && !(nextData.products ?? []).some((p) => current === `category:${p.category ?? ''}`) ? '' : current,
    )
    setBundle(nextData)
    // Prefill from today's sales-derived consumption (same as the original).
    const pre: Record<string, string> = {}
    for (const i of r.data.autoPrefill?.items ?? []) if (i.quantity > 0) pre[i.productCode] = String(i.quantity)
    setQty(pre)
  }, [auth, t])

  useEffect(() => {
    load()
  }, [load])

  const products = bundle?.products ?? []
  const suggestions = useMemo(() => {
    const result: Record<string, number> = {}
    for (const item of bundle?.autoPrefill?.items ?? []) result[item.productCode] = item.quantity
    return result
  }, [bundle])
  const query = search.trim().toLowerCase()
  const categories = useMemo(
    () => Array.from(new Set(products.map((p) => p.category ?? ''))).sort((a, b) => a.localeCompare(b)),
    [products],
  )
  const visible = useMemo(
    () =>
      products.filter(
        (p) =>
          (!query || p.name.toLowerCase().includes(query) || p.productCode.toLowerCase().includes(query)) &&
          (!category || category === `category:${p.category ?? ''}`),
      ),
    [products, query, category],
  )
  const items = useMemo(
    () =>
      Object.entries(qty)
        .map(([productCode, v]) => ({ productCode, quantity: Number(v) || 0 }))
        .filter((i) => i.quantity > 0),
    [qty],
  )

  const save = async () => {
    if (!items.length) return setMsg({ tone: 'error', text: t('cons.noItems') })
    setSaving(true)
    const r = await recordConsumption(items, auth)
    setSaving(false)
    if (r.success) {
      setMsg({ tone: 'info', text: r.data?.message ?? t('cons.saved') })
      setQty({})
      load()
    } else setMsg({ tone: 'error', text: r.error ?? '' })
  }

  const loadHistory = async () => {
    setHistoryLoading(true)
    setHistoryError(null)
    const r = await getConsumptionHistory(historyStart, historyEnd, auth)
    setHistoryLoading(false)
    if (!r.success || !r.data) {
      setHistoryRows([])
      setHistoryError(r.error ?? t('cons.historyFailed'))
      return
    }
    setHistoryRows(r.data.rows)
  }

  const summary = bundle?.summary
  const kf = bundle?.keyfigures
  const [kfSearch, setKfSearch] = useState('')
  const [kfLowOnly, setKfLowOnly] = useState(false)
  const kfItems = kf?.items ?? []
  const kfLow = kfItems.filter((i) => (Number(i.remaining) || 0) <= 0).length
  const kfVisible = useMemo(() => {
    const q = kfSearch.trim().toLowerCase()
    return kfItems.filter(
      (i) =>
        (!q || i.name.toLowerCase().includes(q) || i.productCode.toLowerCase().includes(q) || (i.category ?? '').toLowerCase().includes(q)) &&
        (!kfLowOnly || (Number(i.remaining) || 0) <= (Number(i.reorderQty) || 0)),
    )
  }, [kfItems, kfSearch, kfLowOnly])

  return (
    <Panel title={`🧮 ${t('nav.inventoryConsumption')}`}>
      <SubTabs
        tabs={[
          { id: 'entry', label: `✏️ ${t('cons.entry')}` },
          { id: 'summary', label: `📊 ${t('cons.today')}` },
          { id: 'history', label: `📅 ${t('cons.history')}` },
        ]}
        active={tab}
        onChange={(next) => {
          setTab(next)
          if (next === 'history') void loadHistory()
        }}
      />
      {msg && <Notice tone={msg.tone === 'error' ? 'error' : 'info'}>{msg.text}</Notice>}

      {tab === 'entry' && (
        <>
          <FormRow>
            <Field label={t('checkout.search')}>
              <TextInput type="search" value={search} onChange={(e) => setSearch(e.target.value)} />
            </Field>
            <Field label={t('inv.category')}>
              <Select value={category} onChange={(e) => setCategory(e.target.value)}>
                <option value="">{t('common.all')}</option>
                {categories.map((name) => (
                  <option key={name} value={`category:${name}`}>{name || t('stock.uncategorized')}</option>
                ))}
              </Select>
            </Field>
            <div className="actions" style={{ alignSelf: 'end' }}>
              <Button variant="ghost" onClick={load} disabled={loading}>
                {loading ? t('common.loading') : '↻'}
              </Button>
            </div>
          </FormRow>
          {bundle?.autoPrefill?.totalItems ? (
            <Notice tone="info">
              {t('cons.prefilled')}: {fmtNum(bundle.autoPrefill.totalItems)} ({bundle.autoPrefill.date})
            </Notice>
          ) : null}
          {!!bundle?.autoPrefill?.missingProductCodes?.length && (
            <Notice tone="warn">{t('cons.missingSuggestions')}: {bundle.autoPrefill.missingProductCodes.join(', ')}</Notice>
          )}
          {(bundle?.summary?.totalItems ?? 0) > 0 && <Notice tone="warn">{t('cons.repeatWarning')}</Notice>}
          <DataTable
            rows={visible as unknown as Record<string, unknown>[]}
            rowKey={(r) => String(r.productCode)}
            defaultSort={{ key: 'suggested', dir: 'desc' }}
            empty={loading ? t('common.loading') : t('cons.noProducts')}
            columns={[
              { key: 'name', label: t('purch.product'), sortable: true, render: (r) => <>{String(r.name)}<div className="muted">{String(r.category ?? '')} · {String(r.unit ?? '')}</div></> },
              {
                key: 'suggested',
                label: t('cons.suggested'),
                align: 'right',
                sortable: true,
                sortValue: (r) => suggestions[String(r.productCode)] ?? 0,
                render: (r) => fmtNum(suggestions[String(r.productCode)] ?? 0),
              },
              {
                key: 'qty',
                label: t('cons.consumed'),
                align: 'right',
                sortable: true,
                sortValue: (r) => Number(qty[String(r.productCode)]) || 0,
                render: (r) => (
                  <input
                    type="number"
                    inputMode="decimal"
                    min={0}
                    step="any"
                    style={{ width: '6rem', textAlign: 'right' }}
                    value={qty[String(r.productCode)] ?? ''}
                    onChange={(e) => setQty((q) => ({ ...q, [String(r.productCode)]: e.target.value }))}
                    placeholder="0"
                  />
                ),
              },
            ]}
          />
          <div className="actions">
            <Button onClick={save} disabled={loading || saving || !items.length}>
              {saving ? t('checkout.submitting') : `💾 ${t('cons.save')} (${items.length})`}
            </Button>
          </div>
        </>
      )}

      {tab === 'summary' && (
        <>
          {kf && (
            <Stats
              items={[
                { label: t('stock.products'), value: fmtNum(kfItems.length) },
                { label: t('stock.outOfStock'), value: fmtNum(kfLow), tone: kfLow ? 'bad' : 'ok' },
                { label: t('cons.recordedToday'), value: fmtNum(summary?.totalQuantity ?? 0) },
                { label: t('stock.asOf'), value: kf.historyAsOfDate ? fmtDateOnly(kf.historyAsOfDate) : kf.stockSource || '—' },
              ]}
            />
          )}
          {kf?.needsBootstrap && <Notice tone="warn">{t('stock.historyIncomplete')}</Notice>}
          <FormRow>
            <Field label={t('checkout.search')}>
              <TextInput type="search" value={kfSearch} onChange={(e) => setKfSearch(e.target.value)} />
            </Field>
            <label className="ui-field" style={{ flexDirection: 'row', alignItems: 'center', gap: '0.5rem', alignSelf: 'end' }}>
              <input type="checkbox" checked={kfLowOnly} onChange={(e) => setKfLowOnly(e.target.checked)} /> {t('cons.lowOnly')}
            </label>
            <div className="actions" style={{ alignSelf: 'end' }}>
              <Button variant="ghost" onClick={load} disabled={loading}>
                {loading ? t('common.loading') : '↻'}
              </Button>
            </div>
          </FormRow>
          <DataTable
            rows={kfVisible as unknown as Record<string, unknown>[]}
            rowKey={(r) => String(r.productCode)}
            defaultSort={{ key: 'remaining', dir: 'asc' }}
            empty={loading ? t('common.loading') : t('cons.noSummary')}
            columns={[
              {
                key: 'name',
                label: t('purch.product'),
                sortable: true,
                render: (r) => (
                  <>
                    {String(r.name)}
                    <div className="muted">
                      {String(r.category ?? '')}
                      {r.unit ? ` · ${String(r.unit)}` : ''}
                      {r.supplier ? ` · ${String(r.supplier)}` : ''}
                    </div>
                  </>
                ),
              },
              { key: 'totalPurchased', label: t('stock.purchased'), align: 'right', sortable: true, sortValue: (r) => Number(r.totalPurchased) || 0, render: (r) => fmtNum(Number(r.totalPurchased) || 0) },
              { key: 'totalConsumed', label: t('stock.consumed'), align: 'right', sortable: true, sortValue: (r) => Number(r.totalConsumed) || 0, render: (r) => fmtNum(Number(r.totalConsumed) || 0) },
              { key: 'totalSold', label: t('kitchen.sold'), align: 'right', sortable: true, sortValue: (r) => Number(r.totalSold) || 0, render: (r) => fmtNum(Number(r.totalSold) || 0) },
              {
                key: 'remaining',
                label: t('kitchen.remaining'),
                align: 'right',
                sortable: true,
                sortValue: (r) => Number(r.remaining) || 0,
                render: (r) => {
                  const rem = Number(r.remaining) || 0
                  const reorder = Number(r.reorderQty) || 0
                  return <Badge tone={rem <= 0 ? 'bad' : reorder && rem <= reorder ? 'neutral' : 'ok'}>{fmtNum(rem)}</Badge>
                },
              },
              { key: 'reorderQty', label: t('cons.reorderAt'), align: 'right', sortable: true, sortValue: (r) => Number(r.reorderQty) || 0, render: (r) => (Number(r.reorderQty) ? fmtNum(Number(r.reorderQty)) : '—') },
            ]}
          />

          {summary && summary.items.length > 0 && (
            <>
              <h4 className="muted">📝 {t('cons.recordedToday')}</h4>
              <DataTable
                rows={summary.items as Record<string, unknown>[]}
                rowKey={(r, i) => `${r.productCode ?? i}`}
                columns={[
                  { key: 'name', label: t('purch.product'), sortable: true, sortValue: (r) => String(r.productName ?? r.name ?? r.productCode ?? ''), render: (r) => String(r.productName ?? r.name ?? r.productCode ?? '') },
                  { key: 'unit', label: t('cons.unit'), sortable: true },
                  { key: 'quantity', label: t('cons.consumed'), align: 'right', sortable: true, sortValue: (r) => Number(r.quantity) || 0, render: (r) => fmtNum(Number(r.quantity) || 0) },
                ]}
              />
            </>
          )}
        </>
      )}
      {tab === 'history' && (
            <>
              <FormRow>
                <Field label={t('kitchen.from')}>
                  <TextInput type="date" value={historyStart} disabled={historyLoading} onChange={(e) => {
                    setHistoryStart(e.target.value)
                    setHistoryRows([])
                    setHistoryError(null)
                  }} />
                </Field>
                <Field label={t('kitchen.to')}>
                  <TextInput type="date" value={historyEnd} disabled={historyLoading} onChange={(e) => {
                    setHistoryEnd(e.target.value)
                    setHistoryRows([])
                    setHistoryError(null)
                  }} />
                </Field>
                <div className="actions" style={{ alignSelf: 'end' }}>
                  <Button onClick={loadHistory} disabled={historyLoading}>
                    {historyLoading ? t('common.loading') : t('cons.showHistory')}
                  </Button>
                </div>
              </FormRow>
              {historyError && <Notice tone="error">{historyError}</Notice>}
              <DataTable
                rows={historyRows}
                rowKey={(row) => `${row.date}|${row.productCode}`}
                empty={historyLoading ? t('common.loading') : t('cons.noHistory')}
                columns={[
                  { key: 'date', label: t('kitchen.date'), sortable: true, render: (row) => fmtDateOnly(row.date) },
                  { key: 'name', label: t('purch.product'), sortable: true },
                  { key: 'unit', label: t('cons.unit'), sortable: true },
                  { key: 'quantity', label: t('cons.consumed'), align: 'right', sortable: true, render: (row) => fmtNum(row.quantity) },
                ]}
              />
            </>
      )}
    </Panel>
  )
}
