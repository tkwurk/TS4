import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { apiConfig } from '@/services/api'
import { Badge, Button, Notice, Panel } from '@/components/ui'
import { fetchProducts, productName, type ProductRow } from '@/features/catalog'
import type { Session } from '@/features/auth'
import { useI18n } from '@/i18n'
import { useCart } from '../hooks/useCart'
import { orderNumberOf, productCategory, productImage, productKey, productPrice, productSupplier, productType } from '../types'
import { submitCheckout } from '../api'
import './checkout.css'

const PAYMENTS = ['cash', 'card']

type Msg = { tone: 'info' | 'error'; text: string }

/**
 * Checkout screen — builds a cart from products and submits an order through the
 * Seam 1 API (api.checkout → Orchestrator checkout.completed pipeline). It never
 * touches Sheets directly. The operator identity comes from the signed-in
 * session and is sent as _authContext, matching the saveOrderSecure contract.
 */
export function CheckoutPanel({ session }: { session: Session }) {
  const { t, fmtMoney } = useI18n()
  const configured = Boolean(apiConfig.appsScriptUrl)
  const cart = useCart()

  const [products, setProducts] = useState<ProductRow[]>([])
  const [loadingProducts, setLoadingProducts] = useState(false)
  const [payment, setPayment] = useState(PAYMENTS[0])
  const [supplier, setSupplier] = useState('')
  const [category, setCategory] = useState('')
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

  // Auto-load the menu once when the screen opens.
  const didLoad = useRef(false)
  useEffect(() => {
    if (configured && !didLoad.current) {
      didLoad.current = true
      loadProducts()
    }
  }, [configured, loadProducts])

  // Sellable items grouped Supplier → Category → Products.
  const bySupplier = useMemo(() => {
    const sellable = products.filter((p) => productType(p).toLowerCase() !== 'buy')
    const suppliers = new Map<string, Map<string, ProductRow[]>>()
    for (const p of sellable) {
      const sup = productSupplier(p)
      const cat = productCategory(p)
      if (!suppliers.has(sup)) suppliers.set(sup, new Map())
      const cats = suppliers.get(sup)!
      if (!cats.has(cat)) cats.set(cat, [])
      cats.get(cat)!.push(p)
    }
    return Array.from(suppliers.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([name, cats]) => ({
        supplier: name,
        categories: Array.from(cats.entries())
          .sort((a, b) => a[0].localeCompare(b[0]))
          .map(([category, items]) => ({ category, items })),
      }))
  }, [products])

  // Default the supplier tab to the first available (derived, no effect needed).
  const activeSupplier = bySupplier.find((s) => s.supplier === supplier) ?? bySupplier[0]
  const categories = activeSupplier?.categories ?? []
  const activeCategory = categories.find((c) => c.category === category) ?? categories[0]

  const onSubmit = useCallback(async () => {
    if (!cart.items.length) return setMessage({ tone: 'error', text: t('checkout.cartEmptyError') })
    setSubmitting(true)
    setMessage(null)
    const res = await submitCheckout(cart.items, payment, auth)
    setSubmitting(false)
    if (res.success) {
      const num = orderNumberOf(res.data)
      setMessage({ tone: 'info', text: t('checkout.submitted') + (num ? ` — ${num}` : '') })
      cart.clear()
    } else {
      setMessage({ tone: 'error', text: res.error ?? t('checkout.failed') })
    }
  }, [cart, payment, auth, t])

  return (
    <Panel title={t('checkout.title')}>
      {!configured && (
        <Notice tone="warn">
          <strong>{t('common.setupRequired')}</strong> {t('checkout.setup')}
        </Notice>
      )}

      <div className="checkout-operator">
        <label className="checkout-field">
          {t('checkout.payment')}
          <select value={payment} onChange={(e) => setPayment(e.target.value)}>
            {PAYMENTS.map((p) => (
              <option key={p} value={p}>
                {t('payment.' + p)}
              </option>
            ))}
          </select>
        </label>
        <p className="muted">
          {t('checkout.operator')}: <strong>{session.username}</strong> · {t('checkout.outlet')}:{' '}
          {outlet || t('common.all')}
        </p>
      </div>

      <div className="actions">
        <Button variant="ghost" onClick={loadProducts} disabled={!configured || loadingProducts}>
          {loadingProducts ? t('common.loading') : t('checkout.reloadMenu')}
        </Button>
        <Badge tone="neutral">
          {t('checkout.cart')}: {cart.count}
        </Badge>
      </div>

      {bySupplier.length > 0 && (
        <div className="supplier-nav">
          {bySupplier.map((s) => (
            <button
              key={s.supplier}
              type="button"
              className={'supplier-pill' + (activeSupplier?.supplier === s.supplier ? ' supplier-pill--active' : '')}
              onClick={() => setSupplier(s.supplier)}
            >
              {s.supplier}
            </button>
          ))}
        </div>
      )}

      {categories.length > 0 && (
        <div className="category-nav">
          {categories.map((c) => (
            <button
              key={c.category}
              type="button"
              className={'category-pill' + (activeCategory?.category === c.category ? ' category-pill--active' : '')}
              onClick={() => setCategory(c.category)}
            >
              {c.category}
            </button>
          ))}
        </div>
      )}

      <div className="menu">
        {products.length === 0 ? (
          <p className="muted">{loadingProducts ? t('checkout.loadingMenu') : t('checkout.noProducts')}</p>
        ) : (
          <div className="menu-grid">
            {(activeCategory?.items ?? []).map((p) => {
              const key = productKey(p)
              const name = productName(p)
              const price = productPrice(p)
              const img = productImage(p)
              const qty = cart.qtyOf(key)
              return (
                <div key={key} className="menu-tile">
                  {img && (
                    <img
                      className="menu-tile__img"
                      src={img}
                      alt={name}
                      loading="lazy"
                      referrerPolicy="no-referrer"
                      onError={(e) => {
                        e.currentTarget.style.display = 'none'
                      }}
                    />
                  )}
                  <span className="menu-tile__name">{name}</span>
                  <span className="menu-tile__price">{fmtMoney(price)}</span>
                  <div className="menu-tile__qty">
                    <button
                      type="button"
                      className="qty-btn"
                      onClick={() => cart.setQty(key, qty - 1)}
                      disabled={qty === 0}
                    >
                      −
                    </button>
                    <span className="qty-val">{qty}</span>
                    <button
                      type="button"
                      className="qty-btn"
                      onClick={() => cart.add({ key, name, price })}
                    >
                      +
                    </button>
                  </div>
                </div>
              )
            })}
          </div>
        )}
      </div>

      <h3 className="muted">
        {t('checkout.cart')} ({cart.count})
      </h3>
      {cart.items.length === 0 ? (
        <p className="muted">{t('checkout.cartEmpty')}</p>
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
              <span className="muted">{fmtMoney(i.price * i.qty)}</span>
            </li>
          ))}
        </ul>
      )}

      <div className="checkout-total">
        <span>{t('checkout.total')}</span>
        <span>{fmtMoney(cart.total)}</span>
      </div>

      <div className="actions">
        <Button onClick={onSubmit} disabled={!configured || submitting || cart.items.length === 0}>
          {submitting ? t('checkout.submitting') : t('checkout.submit')}
        </Button>
        {cart.items.length > 0 && (
          <Button variant="ghost" onClick={cart.clear}>
            {t('checkout.clear')}
          </Button>
        )}
      </div>

      {message && <Notice tone={message.tone === 'error' ? 'error' : 'info'}>{message.text}</Notice>}
    </Panel>
  )
}
