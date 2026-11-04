import { useCallback, useEffect, useMemo, useState } from 'react'
import { localDate, localDateTime } from '@/lib/dates'
import type { AuthContext } from '@/services/api'
import { Badge, Button, DataTable, Field, FormRow, Notice, Panel, Stats, SubTabs, TextArea, TextInput } from '@/components/ui'
import { useI18n } from '@/i18n'
import {
  getCookRequests,
  getCookingHistory,
  getCookingProducts,
  getCookingTodayReport,
  submitCookingBatch,
  updateCookRequestStatus,
  type CookRequest,
  type CookingHistoryRow,
  type CookingProduct,
  type TodayReportItem,
} from '../api'

type Tab = 'cook' | 'requests' | 'today' | 'history'
type Msg = { tone: 'info' | 'error'; text: string }

const nowLocal = () => localDateTime()
const today = () => localDate()

/**
 * Kitchen tab — mirrors ui_kitchenCooking.html: log a cooking batch (shared
 * start/end time, qty per cooking product), handle cook requests from sales,
 * today's cooked-vs-sold report, and cooking history.
 */
export function KitchenPanel({ auth }: { auth: AuthContext }) {
  const { t, fmtNum } = useI18n()
  const [tab, setTab] = useState<Tab>('cook')
  const [msg, setMsg] = useState<Msg | null>(null)

  // ── Cook batch ──
  const [products, setProducts] = useState<CookingProduct[]>([])
  const [loadingProducts, setLoadingProducts] = useState(false)
  const [qty, setQty] = useState<Record<string, string>>({})
  const [startTime, setStartTime] = useState(nowLocal)
  const [endTime, setEndTime] = useState(nowLocal)
  const [notes, setNotes] = useState('')
  const [submitting, setSubmitting] = useState(false)

  const loadProducts = useCallback(async () => {
    setLoadingProducts(true)
    const r = await getCookingProducts(auth)
    setLoadingProducts(false)
    if (r.ok) setProducts(r.data?.items ?? [])
    else setMsg({ tone: 'error', text: r.error ?? '' })
  }, [auth])

  const lines = useMemo(
    () =>
      products
        .map((p) => ({ productCode: p.productCode, qty: Number(qty[p.productCode]) || 0 }))
        .filter((l) => l.qty > 0),
    [products, qty],
  )

  const submit = async () => {
    if (!lines.length) return setMsg({ tone: 'error', text: t('kitchen.noLines') })
    setSubmitting(true)
    const r = await submitCookingBatch({ startTime, endTime, notes, lines }, auth)
    setSubmitting(false)
    if (r.ok) {
      setMsg({ tone: 'info', text: `${t('kitchen.saved')} — ${lines.length} ${t('kitchen.lines')}` })
      setQty({})
      setNotes('')
      loadProducts()
      loadRequests()
    } else setMsg({ tone: 'error', text: r.error ?? '' })
  }

  // ── Requests ──
  const [open, setOpen] = useState<CookRequest[]>([])
  const [recentDone, setRecentDone] = useState<CookRequest[]>([])
  const [loadingReq, setLoadingReq] = useState(false)

  const loadRequests = useCallback(async () => {
    setLoadingReq(true)
    const r = await getCookRequests(auth)
    setLoadingReq(false)
    if (r.ok) {
      setOpen(r.data?.open ?? [])
      setRecentDone(r.data?.recentDone ?? [])
    }
  }, [auth])

  const setStatus = async (id: string, status: string) => {
    const r = await updateCookRequestStatus(id, status, auth)
    if (r.ok) loadRequests()
    else setMsg({ tone: 'error', text: r.error ?? '' })
  }

  // ── Today report ──
  const [reportDate, setReportDate] = useState(today)
  const [report, setReport] = useState<TodayReportItem[]>([])
  const [loadingReport, setLoadingReport] = useState(false)

  const loadReport = useCallback(async () => {
    setLoadingReport(true)
    const r = await getCookingTodayReport(reportDate, auth)
    setLoadingReport(false)
    if (r.ok) setReport(r.data?.items ?? [])
    else setMsg({ tone: 'error', text: r.error ?? '' })
  }, [auth, reportDate])

  // ── History ──
  const [hStart, setHStart] = useState(today)
  const [hEnd, setHEnd] = useState(today)
  const [hProduct, setHProduct] = useState('')
  const [history, setHistory] = useState<CookingHistoryRow[]>([])
  const [loadingHist, setLoadingHist] = useState(false)

  const loadHistory = useCallback(async () => {
    setLoadingHist(true)
    const r = await getCookingHistory({ startDate: hStart, endDate: hEnd, productFilter: hProduct }, auth)
    setLoadingHist(false)
    if (r.ok) setHistory(r.data?.rows ?? [])
    else setMsg({ tone: 'error', text: r.error ?? '' })
  }, [auth, hStart, hEnd, hProduct])

  useEffect(() => {
    loadProducts()
    loadRequests()
    const id = setInterval(loadRequests, 60_000)
    return () => clearInterval(id)
  }, [loadProducts, loadRequests])

  const tabs = [
    { id: 'cook' as const, label: `🍳 ${t('kitchen.cook')}` },
    { id: 'requests' as const, label: `🔔 ${t('kitchen.requests')} (${open.length})` },
    { id: 'today' as const, label: `📊 ${t('kitchen.today')}` },
    { id: 'history' as const, label: `📜 ${t('kitchen.history')}` },
  ]

  return (
    <Panel title={`👨‍🍳 ${t('nav.kitchenCooking')}`}>
      <SubTabs tabs={tabs} active={tab} onChange={setTab} />
      {msg && <Notice tone={msg.tone === 'error' ? 'error' : 'info'}>{msg.text}</Notice>}

      {tab === 'cook' && (
        <>
          <FormRow>
            <Field label={t('kitchen.start')}>
              <TextInput type="datetime-local" value={startTime} onChange={(e) => setStartTime(e.target.value)} />
            </Field>
            <Field label={t('kitchen.end')}>
              <TextInput type="datetime-local" value={endTime} onChange={(e) => setEndTime(e.target.value)} />
            </Field>
          </FormRow>
          {loadingProducts ? (
            <p className="muted">{t('common.loading')}</p>
          ) : (
            <DataTable
              rows={products as unknown as Record<string, unknown>[]}
              rowKey={(r) => String(r.productCode)}
              empty={t('kitchen.noProducts')}
              columns={[
                { key: 'name', label: t('kitchen.product') },
                { key: 'cookedToday', label: t('kitchen.cookedToday'), align: 'right', render: (r) => fmtNum(Number(r.cookedToday) || 0) },
                {
                  key: 'qty',
                  label: t('kitchen.qtyCooked'),
                  align: 'right',
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
          )}
          <Field label={t('kitchen.notes')}>
            <TextArea rows={2} value={notes} onChange={(e) => setNotes(e.target.value)} />
          </Field>
          <div className="actions">
            <Button onClick={submit} disabled={submitting || !lines.length}>
              {submitting ? t('checkout.submitting') : `✅ ${t('kitchen.submit')} (${lines.length})`}
            </Button>
            <Button variant="ghost" onClick={loadProducts} disabled={loadingProducts}>
              ↻
            </Button>
          </div>
        </>
      )}

      {tab === 'requests' && (
        <>
          <div className="actions">
            <Button variant="ghost" onClick={loadRequests} disabled={loadingReq}>
              {loadingReq ? t('common.loading') : '↻'}
            </Button>
          </div>
          <h4 className="muted">{t('kitchen.openRequests')}</h4>
          <DataTable
            rows={open as unknown as Record<string, unknown>[]}
            rowKey={(r) => String(r.requestId)}
            empty={t('kitchen.noRequests')}
            columns={[
              { key: 'createdAt', label: t('kitchen.when') },
              { key: 'productName', label: t('kitchen.product') },
              { key: 'requestedQty', label: t('kitchen.qty'), align: 'right' },
              { key: 'requestedBy', label: t('kitchen.by') },
              {
                key: 'status',
                label: t('kitchen.status'),
                render: (r) => <Badge tone={r.status === 'COOKING' ? 'ok' : 'bad'}>{String(r.status)}</Badge>,
              },
              {
                key: 'actions',
                label: '',
                render: (r) => (
                  <span className="actions" style={{ margin: 0 }}>
                    {r.status !== 'COOKING' && (
                      <Button variant="ghost" onClick={() => setStatus(String(r.requestId), 'COOKING')}>
                        🍳 {t('kitchen.startCooking')}
                      </Button>
                    )}
                    <Button onClick={() => setStatus(String(r.requestId), 'DONE')}>✅ {t('kitchen.done')}</Button>
                    <Button variant="ghost" onClick={() => setStatus(String(r.requestId), 'CANCELLED')}>
                      ✕
                    </Button>
                  </span>
                ),
              },
            ]}
          />
          {recentDone.length > 0 && (
            <>
              <h4 className="muted">{t('kitchen.recentDone')}</h4>
              <DataTable
                rows={recentDone as unknown as Record<string, unknown>[]}
                rowKey={(r) => String(r.requestId)}
                columns={[
                  { key: 'updatedAt', label: t('kitchen.when') },
                  { key: 'productName', label: t('kitchen.product') },
                  { key: 'requestedQty', label: t('kitchen.qty'), align: 'right' },
                  { key: 'handledBy', label: t('kitchen.by') },
                ]}
              />
            </>
          )}
        </>
      )}

      {tab === 'today' && (
        <>
          <FormRow>
            <Field label={t('kitchen.date')}>
              <TextInput type="date" value={reportDate} onChange={(e) => setReportDate(e.target.value)} />
            </Field>
            <div className="actions" style={{ alignSelf: 'end' }}>
              <Button onClick={loadReport} disabled={loadingReport}>
                {loadingReport ? t('common.loading') : t('kitchen.loadReport')}
              </Button>
            </div>
          </FormRow>
          {report.length > 0 && (
            <Stats
              items={[
                { label: t('kitchen.products'), value: fmtNum(report.length) },
                { label: t('kitchen.needRecook'), value: fmtNum(report.filter((i) => i.needsRecook).length), tone: report.some((i) => i.needsRecook) ? 'bad' : 'ok' },
                { label: t('kitchen.cookedToday'), value: fmtNum(report.reduce((s, i) => s + (i.cookedToday ?? 0), 0)) },
                { label: t('kitchen.soldToday'), value: fmtNum(report.reduce((s, i) => s + (i.soldToday ?? 0), 0)) },
              ]}
            />
          )}
          <DataTable
            rows={report as unknown as Record<string, unknown>[]}
            rowKey={(r) => String(r.productCode ?? r.name)}
            empty={t('kitchen.noReport')}
            columns={[
              { key: 'name', label: t('kitchen.product') },
              { key: 'leftoverFromYesterday', label: t('kitchen.leftover'), align: 'right', render: (r) => fmtNum(Number(r.leftoverFromYesterday) || 0) },
              { key: 'cookedToday', label: t('kitchen.cooked'), align: 'right', render: (r) => fmtNum(Number(r.cookedToday) || 0) },
              { key: 'soldToday', label: t('kitchen.sold'), align: 'right', render: (r) => fmtNum(Number(r.soldToday) || 0) },
              {
                key: 'remaining',
                label: t('kitchen.remaining'),
                align: 'right',
                render: (r) => (
                  <Badge tone={r.needsRecook ? 'bad' : 'ok'}>
                    {fmtNum(Number(r.remaining) || 0)} / {fmtNum(Number(r.recookThreshold) || 0)}
                  </Badge>
                ),
              },
            ]}
          />
        </>
      )}

      {tab === 'history' && (
        <>
          <FormRow>
            <Field label={t('kitchen.from')}>
              <TextInput type="date" value={hStart} onChange={(e) => setHStart(e.target.value)} />
            </Field>
            <Field label={t('kitchen.to')}>
              <TextInput type="date" value={hEnd} onChange={(e) => setHEnd(e.target.value)} />
            </Field>
            <Field label={t('kitchen.product')}>
              <TextInput value={hProduct} onChange={(e) => setHProduct(e.target.value)} placeholder={t('common.all')} />
            </Field>
            <div className="actions" style={{ alignSelf: 'end' }}>
              <Button onClick={loadHistory} disabled={loadingHist}>
                {loadingHist ? t('common.loading') : t('kitchen.loadHistory')}
              </Button>
            </div>
          </FormRow>
          <DataTable
            rows={history as unknown as Record<string, unknown>[]}
            rowKey={(r, i) => `${r.cookingId}-${i}`}
            empty={t('kitchen.noHistory')}
            columns={[
              { key: 'startTime', label: t('kitchen.start'), render: (r) => (r.startTime ? new Date(String(r.startTime)).toLocaleString() : '') },
              { key: 'productName', label: t('kitchen.product') },
              { key: 'qty', label: t('kitchen.qty'), align: 'right' },
              { key: 'chef', label: t('kitchen.chef') },
              { key: 'notes', label: t('kitchen.notes') },
            ]}
          />
        </>
      )}
    </Panel>
  )
}
