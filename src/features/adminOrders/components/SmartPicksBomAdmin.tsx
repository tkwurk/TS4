import { useCallback, useEffect, useState } from 'react'
import { fmtDateTime } from '@/lib/dates'
import type { AuthContext } from '@/services/api'
import { Badge, Button, DataTable, Field, FormRow, Notice, Panel, Select, Stats, TextInput } from '@/components/ui'
import { useI18n } from '@/i18n'
import {
  fetchBom,
  fetchSmartPicksInfo,
  refreshSmartPicksNow,
  removeSmartPicks,
  scheduleSmartPicks,
  type BomSummary,
  type SmartPicksInfo,
} from '../api'

type Msg = { tone: 'info' | 'error'; text: string }

/** Mirrors ui_adminOrders.html "Smart Picks" pane: cache status + daily trigger management. */
export function SmartPicksAdmin({ auth }: { auth: AuthContext }) {
  const { t, fmtNum } = useI18n()
  const [info, setInfo] = useState<SmartPicksInfo | null>(null)
  const [hour, setHour] = useState('0')
  const [busy, setBusy] = useState(false)
  const [msg, setMsg] = useState<Msg | null>(null)

  const load = useCallback(async () => {
    const r = await fetchSmartPicksInfo(auth)
    if (r.ok) setInfo(r.data)
    else setMsg({ tone: 'error', text: r.error ?? '' })
  }, [auth])
  useEffect(() => {
    load()
  }, [load])

  const run = async (fn: () => Promise<{ ok: boolean; error: string | null; data: unknown }>, okText: (d: unknown) => string) => {
    setBusy(true)
    const r = await fn()
    setBusy(false)
    if (r.ok) {
      setMsg({ tone: 'info', text: okText(r.data) })
      load()
    } else setMsg({ tone: 'error', text: r.error ?? '' })
  }

  return (
    <Panel title={`⚡ ${t('adm.smartPicks')}`}>
      {msg && <Notice tone={msg.tone}>{msg.text}</Notice>}
      <p className="muted">{t('adm.smartPicksHelp')}</p>
      {info && (
        <Stats
          items={[
            { label: t('adm.trigger'), value: info.triggerActive ? t('adm.active') : t('adm.inactive'), tone: info.triggerActive ? 'ok' : 'bad' },
            { label: t('adm.lastUpdated'), value: info.lastUpdated ? fmtDateTime(info.lastUpdated) : '—' },
            { label: t('adm.cachedRows'), value: fmtNum(Number(info.totalRows) || 0) },
          ]}
        />
      )}
      <FormRow>
        <Field label={t('adm.hour')}>
          <Select value={hour} onChange={(e) => setHour(e.target.value)}>
            {Array.from({ length: 24 }, (_, h) => (
              <option key={h} value={h}>
                {String(h).padStart(2, '0')}:00
              </option>
            ))}
          </Select>
        </Field>
        <div className="actions" style={{ alignSelf: 'end' }}>
          <Button disabled={busy} onClick={() => run(() => scheduleSmartPicks(Number(hour), auth), (d) => (d as { message?: string })?.message ?? t('common.saved'))}>
            ⏰ {t('adm.schedule')}
          </Button>
          <Button variant="ghost" disabled={busy} onClick={() => run(() => removeSmartPicks(auth), () => t('adm.triggerRemoved'))}>
            🗑️ {t('adm.removeTrigger')}
          </Button>
          <Button disabled={busy} onClick={() => run(() => refreshSmartPicksNow(auth), (d) => (d as { message?: string })?.message ?? t('adm.refreshed'))}>
            {busy ? t('common.loading') : `🔄 ${t('adm.refreshNow')}`}
          </Button>
        </div>
      </FormRow>
    </Panel>
  )
}

/** Mirrors ui_adminOrders.html "BOM consumption" pane: read-only recipe summary. */
export function BomAdmin({ auth }: { auth: AuthContext }) {
  const { t, fmtNum } = useI18n()
  const [bom, setBom] = useState<BomSummary | null>(null)
  const [q, setQ] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const load = useCallback(async () => {
    setLoading(true)
    const r = await fetchBom(auth)
    setLoading(false)
    if (r.ok) setBom(r.data)
    else setError(r.error)
  }, [auth])
  useEffect(() => {
    load()
  }, [load])

  const ql = q.trim().toLowerCase()
  const lines = (bom?.lines ?? []).filter((l) => !ql || l.unitProductCode.toLowerCase().includes(ql) || l.recipeItemCode.toLowerCase().includes(ql) || l.bomId.toLowerCase().includes(ql))

  return (
    <Panel title={`🧾 ${t('adm.bom')}`}>
      {error && <Notice tone="error">{error}</Notice>}
      <p className="muted">{t('adm.bomHelp')}</p>
      {bom && (
        <Stats
          items={[
            { label: t('adm.bomLines'), value: fmtNum(Number(bom.totalLines) || bom.lines.length) },
            { label: t('adm.unitProducts'), value: fmtNum(Number(bom.unitProductCount) || 0) },
            { label: t('adm.sheet'), value: bom.sheetName ?? '—' },
          ]}
        />
      )}
      <FormRow>
        <Field label={t('checkout.search')}>
          <TextInput type="search" value={q} onChange={(e) => setQ(e.target.value)} />
        </Field>
        <div className="actions" style={{ alignSelf: 'end' }}>
          <Button variant="ghost" onClick={load} disabled={loading}>
            {loading ? t('common.loading') : '↻'}
          </Button>
        </div>
      </FormRow>
      <DataTable
        rows={lines as unknown as Record<string, unknown>[]}
        rowKey={(r, i) => `${r.bomId}-${i}`}
        empty={loading ? t('common.loading') : t('adm.noBom')}
        columns={[
          { key: 'bomId', label: 'BOM', render: (r) => <code>{String(r.bomId)}</code> },
          { key: 'unitProductCode', label: t('adm.unitProduct') },
          { key: 'recipeItemCode', label: t('adm.recipeItem') },
          { key: 'itemQuantity', label: t('purch.qty'), align: 'right', render: (r) => String(r.itemQuantityDisplay ?? r.itemQuantity ?? '') },
          { key: 'status', label: t('kitchen.status'), render: (r) => <Badge tone={String(r.status ?? 'Active').toLowerCase() === 'active' ? 'ok' : 'neutral'}>{String(r.status ?? '')}</Badge> },
        ]}
      />
    </Panel>
  )
}
