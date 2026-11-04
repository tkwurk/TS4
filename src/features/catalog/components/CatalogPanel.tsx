import { useCallback, useState } from 'react'
import { api, apiConfig } from '@/services/api'
import { Badge, Button, Notice, Panel } from '@/components/ui'
import { fetchProducts } from '../api'
import { productName, type ProductRow } from '../types'

type Health = 'unknown' | 'ok' | 'down'

/**
 * Minimal end-to-end demo of the architecture: this panel reaches the Google
 * Sheets data ONLY through the Seam 1 JSON API → Orchestrator → Inventory
 * Worker → Repo. It never touches the Sheets API directly.
 */
export function CatalogPanel() {
  const configured = Boolean(apiConfig.appsScriptUrl)
  const [health, setHealth] = useState<Health>('unknown')
  const [products, setProducts] = useState<ProductRow[]>([])
  const [count, setCount] = useState<number | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const checkHealth = useCallback(async () => {
    const res = await api.health()
    setHealth(res.success ? 'ok' : 'down')
  }, [])

  const loadProducts = useCallback(async () => {
    setLoading(true)
    setError(null)
    const res = await fetchProducts({})
    setLoading(false)
    if (res.success && res.data) {
      setProducts(res.data.products ?? [])
      setCount(res.data.count ?? res.data.products?.length ?? 0)
    } else {
      setError(res.error ?? 'Unknown error')
    }
  }, [])

  const healthTone = health === 'ok' ? 'ok' : health === 'down' ? 'bad' : 'neutral'
  const healthLabel = health === 'ok' ? 'OK' : health === 'down' ? 'Down' : '—'

  return (
    <Panel title="Backend API (Seam 1)">
      {!configured && (
        <Notice tone="warn">
          <strong>Setup required.</strong> Set <code>VITE_APPS_SCRIPT_URL</code> in{' '}
          <code>.env.local</code> to your deployed Apps Script <code>/exec</code> URL.
        </Notice>
      )}

      <div className="actions">
        <Button variant="ghost" onClick={checkHealth} disabled={!configured}>
          Check health
        </Button>
        <Button onClick={loadProducts} disabled={!configured || loading}>
          {loading ? 'Loading…' : 'Load products'}
        </Button>
      </div>

      <p className="status">
        Health: <Badge tone={healthTone}>{healthLabel}</Badge>
        {count !== null && <> · Products: {count}</>}
      </p>

      {error && <Notice tone="error">{error}</Notice>}

      {products.length > 0 && (
        <ul className="product-list">
          {products.slice(0, 25).map((p, i) => (
            <li key={i}>{productName(p)}</li>
          ))}
        </ul>
      )}
    </Panel>
  )
}
