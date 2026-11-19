import { useCallback, useEffect, useMemo, useState } from 'react'
import { fmtDateOnly } from '@/lib/dates'
import { api } from '@/services/api'
import type { AuthContext } from '@/services/api'
import { Badge, Button, DataTable, Field, FormRow, Notice, Panel, Select, Stats, TextInput } from '@/components/ui'
import { useI18n } from '@/i18n'

interface StockItem extends Record<string, unknown> {
  productCode: string
  name: string
  unit: string
  category: string
  purchased: number
  consumed: number
  sold: number
  remaining: number
}
interface StockResult {
  outlet: string
  items: StockItem[]
  stockSource: string
  historyAsOfDate: string
  todayDate: string
  historyIncomplete: boolean
}

const fmtDate = (v: string) => (v ? fmtDateOnly(v) : '—')

/** Stock — current ledger totals (purchased − consumed − sold) per buy product for the outlet. */
export function StockPanel({ auth }: { auth: AuthContext }) {
  const { t, fmtNum } = useI18n()
  const [data, setData] = useState<StockResult | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [search, setSearch] = useState('')
  const [category, setCategory] = useState('')
  const [stockStatus, setStockStatus] = useState('all')

  const load = useCallback(async () => {
    setLoading(true)
    setError(null)
    const r = await api.stockReport<StockResult>(auth)
    setLoading(false)
    if (!r.success || !r.data) {
      setData(null)
      return setError(r.error || 'Stock report is unavailable.')
    }
    const nextData = r.data
    setCategory((current) => current && !nextData.items.some((item) => current === `category:${item.category}`) ? '' : current)
    setData(nextData)
  }, [auth])

  useEffect(() => {
    load()
  }, [load])

  const q = search.trim().toLowerCase()
  const categories = useMemo(
    () => Array.from(new Set((data?.items ?? []).map((item) => item.category))).sort((a, b) => a.localeCompare(b)),
    [data],
  )
  const items = useMemo(
    () =>
      (data?.items ?? []).filter(
        (i) =>
          (!q || i.name.toLowerCase().includes(q) || i.productCode.toLowerCase().includes(q)) &&
          (!category || category === `category:${i.category}`) &&
          (stockStatus === 'all' || (stockStatus === 'in' ? i.remaining > 0 : i.remaining <= 0)),
      ),
    [data, q, category, stockStatus],
  )
  const low = (data?.items ?? []).filter((i) => i.remaining <= 0).length

  return (
    <Panel title={`📊 ${t('nav.stockReport')}`}>
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
        <Field label={t('stock.status')}>
          <Select value={stockStatus} onChange={(e) => setStockStatus(e.target.value)}>
            <option value="all">{t('common.all')}</option>
            <option value="in">{t('stock.inStock')}</option>
            <option value="out">{t('stock.outOfStock')}</option>
          </Select>
        </Field>
        <div className="actions" style={{ alignSelf: 'end' }}>
          <Button variant="ghost" onClick={load} disabled={loading}>
            {loading ? t('common.loading') : '↻'}
          </Button>
        </div>
      </FormRow>
      {error && <Notice tone="error">{error}</Notice>}
      {data?.historyIncomplete && <Notice tone="warn">{t('stock.historyIncomplete')}</Notice>}
      {data && (
        <Stats
          items={[
            { label: t('stock.products'), value: fmtNum(data.items.length) },
            { label: t('stock.outOfStock'), value: fmtNum(low), tone: low ? 'bad' : 'ok' },
            { label: t('stock.asOf'), value: fmtDate(data.todayDate || data.historyAsOfDate) },
          ]}
        />
      )}
      {data && <p className="muted">{t('stock.sortHint')}</p>}
      <DataTable
        rows={items}
        rowKey={(r) => r.productCode}
        defaultSort={{ key: 'remaining', dir: 'asc' }}
        empty={loading ? t('common.loading') : t('stock.none')}
        columns={[
          { key: 'name', label: t('purch.product'), sortable: true, render: (r) => <>{r.name}<div className="muted">{r.category} · {r.unit}</div></> },
          { key: 'purchased', label: t('stock.purchased'), align: 'right', sortable: true, render: (r) => fmtNum(r.purchased) },
          { key: 'consumed', label: t('stock.consumed'), align: 'right', sortable: true, render: (r) => fmtNum(r.consumed) },
          { key: 'sold', label: t('kitchen.sold'), align: 'right', sortable: true, render: (r) => fmtNum(r.sold) },
          {
            key: 'remaining',
            label: t('kitchen.remaining'),
            align: 'right',
            sortable: true,
            render: (r) => <Badge tone={r.remaining <= 0 ? 'bad' : 'ok'}>{fmtNum(r.remaining)}</Badge>,
          },
        ]}
      />
    </Panel>
  )
}
