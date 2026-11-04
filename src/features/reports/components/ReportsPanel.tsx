import { useCallback, useState } from 'react'
import { apiConfig } from '@/services/api'
import type { AuthContext } from '@/services/api'
import { Button, Notice, Panel } from '@/components/ui'
import { fetchSalesSummary } from '../api'
import { summaryOf, type SalesSummary } from '../types'

const money = (n: number) => '$' + n.toFixed(2)

/**
 * Sales summary — reads the Orders sheet through the Reporting Worker's
 * Repo-backed path (Seam 2) via the Seam 1 API. Demonstrates the reporting port.
 */
export function ReportsPanel({ auth }: { auth?: AuthContext }) {
  const configured = Boolean(apiConfig.appsScriptUrl)
  const [summary, setSummary] = useState<SalesSummary | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const load = useCallback(async () => {
    setLoading(true)
    setError(null)
    const res = await fetchSalesSummary(auth?.outlet ?? '', auth)
    setLoading(false)
    if (res.success) {
      setSummary(summaryOf(res.data) ?? null)
    } else {
      setError(res.error ?? 'Failed to load summary')
    }
  }, [auth])

  return (
    <Panel title="Sales summary (Repo-backed)">
      <div className="actions">
        <Button onClick={load} disabled={!configured || loading}>
          {loading ? 'Loading…' : 'Load sales summary'}
        </Button>
      </div>

      {error && <Notice tone="error">{error}</Notice>}

      {summary && (
        <ul className="product-list">
          <li>Outlet: {summary.outlet ?? 'all'}</li>
          <li>Order lines: {summary.lineCount}</li>
          <li>Total sales: {money(summary.totalSales)}</li>
          <li>Items sold: {summary.totalQty}</li>
          {Object.entries(summary.byPayment).map(([method, amount]) => (
            <li key={method}>
              {method}: {money(amount)}
            </li>
          ))}
        </ul>
      )}
    </Panel>
  )
}
