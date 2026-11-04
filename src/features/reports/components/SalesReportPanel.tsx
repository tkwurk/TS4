import { useCallback, useEffect, useMemo, useState } from 'react'
import type { AuthContext } from '@/services/api'
import { Button, DataTable, Field, FormRow, Notice, Panel, Stats, SubTabs, TextInput } from '@/components/ui'
import { useI18n } from '@/i18n'
import { fetchOrderLines, fetchSalesSummary, type OrderLine, type SalesSummary } from '../reportsApi'

type View = 'summary' | 'products' | 'lines'
const today = () => new Date().toISOString().slice(0, 10)

/** Daily sales — mirrors ui_showSalesSummary.html / ui_index.html (getSalesSummaryByDate + getOrderLinesByDate). */
export function SalesReportPanel({ auth }: { auth: AuthContext }) {
  const { t, fmtMoney, fmtNum } = useI18n()
  const [date, setDate] = useState(today)
  const [view, setView] = useState<View>('summary')
  const [summary, setSummary] = useState<SalesSummary | null>(null)
  const [lines, setLines] = useState<OrderLine[]>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const load = useCallback(async () => {
    setLoading(true)
    setError(null)
    const [s, l] = await Promise.all([fetchSalesSummary(date, auth), fetchOrderLines(date, auth)])
    setLoading(false)
    if (!s) setError(t('sales.failed'))
    setSummary(s)
    setLines(l)
  }, [auth, date, t])

  useEffect(() => {
    load()
  }, [load])

  const products = useMemo(() => {
    const s = summary?.summary ?? {}
    return Object.entries(s)
      .map(([name, v]) => ({ name, qty: Number(v?.qty) || 0, total: Number(v?.total) || 0, category: String(v?.category ?? '') }))
      .sort((a, b) => b.total - a.total)
  }, [summary])

  const shift = (days: number) => {
    const d = new Date(date)
    d.setDate(d.getDate() + days)
    setDate(d.toISOString().slice(0, 10))
  }

  return (
    <Panel title={`📅 ${t('nav.monthlySalesReport')}`}>
      <FormRow>
        <Field label={t('kitchen.date')}>
          <TextInput type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </Field>
        <div className="actions" style={{ alignSelf: 'end' }}>
          <Button variant="ghost" onClick={() => shift(-1)}>‹</Button>
          <Button variant="ghost" onClick={() => setDate(today())}>{t('sales.today')}</Button>
          <Button variant="ghost" onClick={() => shift(1)}>›</Button>
          <Button variant="ghost" onClick={load} disabled={loading}>
            {loading ? t('common.loading') : '↻'}
          </Button>
        </div>
      </FormRow>
      {error && <Notice tone="error">{error}</Notice>}

      {summary && (
        <Stats
          items={[
            { label: t('dep.totalSales'), value: fmtMoney(summary.totalSales), tone: 'ok' },
            { label: t('sales.orders'), value: fmtNum(summary.ordersCount) },
            { label: t('dep.cashSales'), value: fmtMoney(summary.totalSalesByCash) },
            { label: t('dep.nonCashSales'), value: fmtMoney(summary.totalSalesByNonCash) },
            { label: t('dep.totalDeposits'), value: fmtMoney(summary.totalDeposits) },
            { label: t('sales.missed'), value: `${fmtNum(summary.missedOrdersCount)} · ${fmtMoney(summary.missedOrdersTotal)}`, tone: summary.missedOrdersCount ? 'bad' : undefined },
          ]}
        />
      )}

      <SubTabs
        tabs={[
          { id: 'summary', label: t('sales.byPayment') },
          { id: 'products', label: t('sales.byProduct') },
          { id: 'lines', label: `${t('sales.lines')} (${lines.length})` },
        ]}
        active={view}
        onChange={setView}
      />

      {view === 'summary' && (
        <DataTable
          rows={Object.entries(summary?.salesByPaymentMethod ?? {}).map(([k, v]) => ({ k, v }))}
          rowKey={(r) => String(r.k)}
          empty={loading ? t('common.loading') : t('sales.none')}
          columns={[
            { key: 'k', label: t('checkout.payment') },
            { key: 'v', label: t('inv.total'), align: 'right', render: (r) => fmtMoney(Number(r.v) || 0) },
          ]}
        />
      )}
      {view === 'products' && (
        <DataTable
          rows={products as unknown as Record<string, unknown>[]}
          rowKey={(r) => String(r.name)}
          empty={loading ? t('common.loading') : t('sales.none')}
          columns={[
            { key: 'name', label: t('purch.product') },
            { key: 'category', label: t('inv.category') },
            { key: 'qty', label: t('purch.qty'), align: 'right', render: (r) => fmtNum(Number(r.qty) || 0) },
            { key: 'total', label: t('inv.total'), align: 'right', render: (r) => fmtMoney(Number(r.total) || 0) },
          ]}
        />
      )}
      {view === 'lines' && (
        <DataTable
          rows={lines as unknown as Record<string, unknown>[]}
          rowKey={(_, i) => String(i)}
          empty={loading ? t('common.loading') : t('sales.none')}
          columns={[
            { key: 'time', label: t('kitchen.when'), render: (r) => String(r.time ?? r.timestamp ?? '') },
            { key: 'orderNumber', label: t('orders.col.order') },
            { key: 'product', label: t('purch.product'), render: (r) => String(r.product ?? r.name ?? '') },
            { key: 'qty', label: t('purch.qty'), align: 'right' },
            { key: 'total', label: t('inv.total'), align: 'right', render: (r) => fmtMoney(Number(r.total) || 0) },
            { key: 'payment', label: t('checkout.payment') },
            { key: 'username', label: t('kitchen.by') },
          ]}
        />
      )}
    </Panel>
  )
}
