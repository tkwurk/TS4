import { useCallback, useEffect, useMemo, useState } from 'react'
import { apiConfig } from '@/services/api'
import { Badge, Button, Notice, Panel } from '@/components/ui'
import { fetchProducts, productName, type ProductRow } from '@/features/catalog'
import type { Session } from '@/features/auth'
import { useCart } from '../hooks/useCart'
import { orderNumberOf, productKey, productPrice } from '../types'
import { submitCheckout } from '../api'
import './checkout.css'

const money = (n: number) => '$' + n.toFixed(2)
const PAYMENTS = ['cash', 'card']

type Msg = { tone: 'info' | 'error'; text: string }

/**
 * Checkout screen — builds a cart from products and submits an order through the
 * Seam 1 API (api.checkout → Orchestrator checkout.completed pipeline). It never
 * touches Sheets directly. The operator identity comes from the signed-in
 * session and is sent as _authContext, matching the saveOrderSecure contract.
 */
export function CheckoutPanel({ session }: { session: Session }) {
  const configured = Boolean(apiConfig.appsScriptUrl)
  const cart = useCart()

  const [products, setProducts] = useState<ProductRow[]>([])
  const [loadingProducts, setLoadingProducts] = useState(false)
  const [payment, setPayment] = useState(PAYMENTS[0])
  const [submitting, setSubmitting] = useState(false)
  const [message, setMessage] = useState<Msg | null>(null)

  const outlet = session.outlet
  const { clear: clearCart } = cart
  const auth = useMemo(
    () => ({ username: session.username, deviceId: session.deviceId, outlet }),
    [session.username, session.deviceId, outlet],
  )

  const loadProducts = useCallback(async () => {
    setLoadingProducts(true)
    setMessage(null)
    const res = await fetchProducts(outlet ? { outlet } : {})
    setLoadingProducts(false)
    if (res.success && res.data) setProducts(res.data.products ?? [])
    else setMessage({ tone: 'error', text: res.error ?? 'Failed to load products' })
  }, [outlet])

  useEffect(() => {
    clearCart()
  }, [session.username, clearCart])

  const onSubmit = useCallback(async () => {
    if (!cart.items.length) return setMessage({ tone: 'error', text: 'Cart is empty.' })
    setSubmitting(true)
    setMessage(null)
    const res = await submitCheckout(cart.items, payment, auth)
    setSubmitting(false)
    if (res.success) {
      const num = orderNumberOf(res.data)
      setMessage({ tone: 'info', text: 'Order submitted ✅' + (num ? ` — ${num}` : '') })
      cart.clear()
    } else {
      setMessage({ tone: 'error', text: res.error ?? 'Checkout failed.' })
    }
  }, [cart, payment, auth])

  return (
    <Panel title="Checkout">
      {!configured && (
        <Notice tone="warn">
          <strong>Setup required.</strong> Set <code>VITE_APPS_SCRIPT_URL</code> in{' '}
          <code>.env.local</code> to submit orders.
        </Notice>
      )}

      <div className="checkout-operator">
        <label className="checkout-field">
          Payment
          <select value={payment} onChange={(e) => setPayment(e.target.value)}>
            {PAYMENTS.map((p) => (
              <option key={p} value={p}>
                {p}
              </option>
            ))}
          </select>
        </label>
        <p className="muted">
          Operator: <strong>{session.username}</strong> · Outlet: {outlet || '(all)'}
        </p>
      </div>

      <div className="actions">
        <Button variant="ghost" onClick={loadProducts} disabled={!configured || loadingProducts}>
          {loadingProducts ? 'Loading…' : 'Load products'}
        </Button>
        <Badge tone="neutral">Cart: {cart.count}</Badge>
      </div>

      <div className="checkout-cols">
        <div>
          <h3 className="muted">Products</h3>
          {products.length === 0 ? (
            <p className="muted">No products loaded.</p>
          ) : (
            <ul className="picker">
              {products.slice(0, 50).map((p) => {
                const price = productPrice(p)
                return (
                  <li key={productKey(p)}>
                    <span className="name">{productName(p)}</span>
                    <span className="muted">{money(price)}</span>
                    <Button
                      onClick={() =>
                        cart.add({ key: productKey(p), name: productName(p), price })
                      }
                    >
                      Add
                    </Button>
                  </li>
                )
              })}
            </ul>
          )}
        </div>

        <div>
          <h3 className="muted">Cart</h3>
          {cart.items.length === 0 ? (
            <p className="muted">Cart is empty.</p>
          ) : (
            <ul className="cart">
              {cart.items.map((i) => (
                <li key={i.key}>
                  <span className="name">{i.name}</span>
                  <span className="qty">
                    <button type="button" onClick={() => cart.setQty(i.key, i.qty - 1)}>
                      −
                    </button>
                    {i.qty}
                    <button type="button" onClick={() => cart.setQty(i.key, i.qty + 1)}>
                      +
                    </button>
                  </span>
                  <span className="muted">{money(i.price * i.qty)}</span>
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>

      <div className="checkout-total">
        <span>Total</span>
        <span>{money(cart.total)}</span>
      </div>

      <div className="actions">
        <Button onClick={onSubmit} disabled={!configured || submitting || cart.items.length === 0}>
          {submitting ? 'Submitting…' : 'Submit order'}
        </Button>
        {cart.items.length > 0 && (
          <Button variant="ghost" onClick={cart.clear}>
            Clear
          </Button>
        )}
      </div>

      {message && <Notice tone={message.tone === 'error' ? 'error' : 'info'}>{message.text}</Notice>}
    </Panel>
  )
}
