import { useCallback, useEffect, useMemo, useState } from 'react'
import type { AuthContext } from '@/services/api'
import { Badge, Button, DataTable, Field, FormRow, Notice, Panel, Select, TextInput } from '@/components/ui'
import { useI18n } from '@/i18n'
import { fetchPaymentMethodsForModule, paymentIcon } from '@/features/checkout'
import type { PaymentMethod } from '@/features/checkout'
import {
  fetchAllProducts,
  fetchExpenseAccounts,
  fetchOrderedPurchases,
  isActive,
  markReceived,
  outletMatches,
  recordPurchase,
  type Account,
  type BuyProduct,
  type OrderedPurchase,
  type PurchaseItem,
} from '../api'
import './purchasing.css'

export type PurchasingMode = 'normal' | 'local' | 'receive'
type Step = 'items' | 'payment'
type Msg = { tone: 'info' | 'error'; text: string }

const str = (v: unknown) => (v == null ? '' : String(v))
const num = (v: unknown) => Number(v) || 0

/**
 * Purchasing / Local buy / Receive — mirrors ui_purchasing.html + ui_purchasingLocal.html:
 * Buy-type products (MODE≠LO for normal; MODE L/LO + misc items for local), expense
 * account, qty/price per line (variable items adjustable), expiry date, payment
 * method (purchasing module) → recordPurchaseSecure. Receive lists ORDERED rows
 * and marks them RECEIVED with the received qty.
 */
export function PurchasingPanel({ auth, mode }: { auth: AuthContext; mode: PurchasingMode }) {
  const { t, fmtMoney, fmtNum } = useI18n()
  const [msg, setMsg] = useState<Msg | null>(null)

  if (mode === 'receive') return <ReceivePanel auth={auth} />

  return <BuyPanel auth={auth} local={mode === 'local'} msg={msg} setMsg={setMsg} t={t} fmtMoney={fmtMoney} fmtNum={fmtNum} />
}

function BuyPanel({
  auth,
  local,
  msg,
  setMsg,
  t,
  fmtMoney,
  fmtNum,
}: {
  auth: AuthContext
  local: boolean
  msg: Msg | null
  setMsg: (m: Msg | null) => void
  t: (k: string) => string
  fmtMoney: (n: number) => string
  fmtNum: (n: number) => string
}) {
  const [products, setProducts] = useState<BuyProduct[]>([])
  const [accounts, setAccounts] = useState<Account[]>([])
  const [account, setAccount] = useState('')
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const [category, setCategory] = useState('')
  const [cart, setCart] = useState<PurchaseItem[]>([])
  const [step, setStep] = useState<Step>('items')
  const [methods, setMethods] = useState<PaymentMethod[]>([])
  const [submitting, setSubmitting] = useState(false)
  const [miscName, setMiscName] = useState('')
  const [miscPrice, setMiscPrice] = useState('')

  const load = useCallback(async () => {
    setLoading(true)
    const [all, accs] = await Promise.all([fetchAllProducts(auth), fetchExpenseAccounts()])
    setLoading(false)
    const buy = all.filter((p) => {
      if (!isActive(p) || str(p.TYPE) !== 'Buy') return false
      const m = str(p.MODE).trim().toUpperCase()
      const modeOk = local ? m === 'L' || m === 'LO' : m !== 'LO'
      return modeOk && outletMatches(str(p.OUTLET), auth.outlet ?? '')
    })
    setProducts(buy)
    setAccounts(accs)
    if (accs.length && !account) setAccount(accs[0].name)
  }, [auth, local, account])

  useEffect(() => {
    load()
  }, [load])

  const categories = useMemo(() => Array.from(new Set(products.map((p) => str(p.CATEGORY) || 'Other'))).sort(), [products])
  const query = search.trim().toLowerCase()
  const visible = useMemo(
    () =>
      products.filter((p) => {
        if (query) return str(p.NAME).toLowerCase().includes(query)
        return !category || (str(p.CATEGORY) || 'Other') === category
      }),
    [products, query, category],
  )

  const keyOf = (p: BuyProduct) => str(p.PRODUCTCODE) || `${str(p.CATEGORY)}::${str(p.NAME)}`
  const toItem = (p: BuyProduct): Omit<PurchaseItem, 'qty'> => ({
    key: keyOf(p),
    name: str(p.NAME),
    productName: str(p.NAME),
    code: str(p.PRODUCTCODE),
    category: str(p.CATEGORY),
    supplier: str(p.SUPPLIER),
    buyUnit: str(p.BUYUNIT),
    buyUnitSize: str(p.BUYUNITSIZE),
    price: num(p.BUYPRICE) || num(p.PRICE),
    mode: local ? str(p.MODE) || 'L' : str(p.MODE),
    isVariable: str(p.QUANTITYVARIABLE).toLowerCase() === 'yes',
  })

  const lineOf = (key: string) => cart.find((c) => c.key === key)
  const setLine = (key: string, patch: Partial<PurchaseItem>, base?: Omit<PurchaseItem, 'qty'>) =>
    setCart((prev) => {
      const i = prev.findIndex((c) => c.key === key)
      if (i >= 0) {
        const next = [...prev]
        next[i] = { ...next[i], ...patch }
        return next.filter((c) => c.qty > 0)
      }
      if (!base) return prev
      const item = { ...base, qty: 0, ...patch } as PurchaseItem
      return item.qty > 0 ? [...prev, item] : prev
    })

  const addMisc = () => {
    const name = miscName.trim()
    const price = parseFloat(miscPrice) || 0
    if (name.length < 2) return setMsg({ tone: 'error', text: t('custom.invalidName') })
    if (price <= 0) return setMsg({ tone: 'error', text: t('custom.invalidPrice') })
    const key = `misc::${name.toLowerCase()}::${price}`
    setLine(key, { qty: (lineOf(key)?.qty ?? 0) + 1 }, {
      key, name, productName: name, code: '', category: 'Local Misc', supplier: '', buyUnit: '', buyUnitSize: '', price, mode: 'L', isMisc: true,
    })
    setMiscName('')
    setMiscPrice('')
  }

  const total = cart.reduce((s, c) => s + c.price * c.qty, 0)

  const proceed = async () => {
    if (!account) return setMsg({ tone: 'error', text: t('purch.noAccount') })
    if (!cart.length) return setMsg({ tone: 'error', text: t('purch.noItems') })
    setMsg(null)
    setStep('payment')
    if (!methods.length) setMethods(await fetchPaymentMethodsForModule(auth, local ? 'localpurchasing' : 'purchasing'))
  }

  const pay = async (method: string) => {
    setSubmitting(true)
    const r = await recordPurchase({ items: cart, account, payment: method }, auth)
    setSubmitting(false)
    if (r.ok) {
      setMsg({ tone: 'info', text: `${t('purch.saved')} — ${fmtMoney(total)}` })
      setCart([])
      setStep('items')
    } else setMsg({ tone: 'error', text: r.error ?? '' })
  }

  return (
    <Panel title={local ? `🏪 ${t('nav.purchasingLocal')}` : `📦 ${t('nav.purchasing')}`}>
      {msg && <Notice tone={msg.tone === 'error' ? 'error' : 'info'}>{msg.text}</Notice>}

      {step === 'payment' ? (
        <div className="purch-pay">
          <div className="actions">
            <Button variant="ghost" onClick={() => setStep('items')} disabled={submitting}>
              ‹ {t('pay.back')}
            </Button>
            <strong>
              {t('checkout.total')}: {fmtMoney(total)} · {t('purch.account')}: {account}
            </strong>
          </div>
          <h4>💳 {t('pay.method')}</h4>
          {methods.length === 0 ? (
            <Notice tone="warn">{t('purch.noMethods')}</Notice>
          ) : (
            <div className="pay__methods">
              {methods.map((m) => (
                <button key={m.pn || m.name} type="button" className="pay__method" onClick={() => pay(m.name)} disabled={submitting}>
                  <span className="pay__method-icon">{paymentIcon(m.name)}</span>
                  <span>{m.name}</span>
                </button>
              ))}
            </div>
          )}
        </div>
      ) : (
        <div className="purch">
          <div className="purch__main">
            <FormRow>
              <Field label={t('purch.account')}>
                <Select value={account} onChange={(e) => setAccount(e.target.value)}>
                  {accounts.length === 0 && <option value="">—</option>}
                  {accounts.map((a) => (
                    <option key={a.name} value={a.name}>
                      {a.name}
                    </option>
                  ))}
                </Select>
              </Field>
              <Field label={t('checkout.search')}>
                <TextInput type="search" value={search} onChange={(e) => setSearch(e.target.value)} placeholder={t('checkout.search')} />
              </Field>
            </FormRow>

            {!query && (
              <div className="category-nav">
                <button type="button" className={'category-pill' + (!category ? ' category-pill--active' : '')} onClick={() => setCategory('')}>
                  {t('common.all')}
                </button>
                {categories.map((c) => (
                  <button key={c} type="button" className={'category-pill' + (category === c ? ' category-pill--active' : '')} onClick={() => setCategory(c)}>
                    {c}
                  </button>
                ))}
              </div>
            )}

            {loading ? (
              <p className="muted">{t('common.loading')}</p>
            ) : (
              <DataTable
                rows={visible as Record<string, unknown>[]}
                rowKey={(r) => keyOf(r as BuyProduct)}
                empty={t('purch.noProducts')}
                columns={[
                  { key: 'NAME', label: t('purch.product'), render: (r) => <>{str(r.NAME)}<div className="muted">{str(r.SUPPLIER)} · {str(r.BUYUNIT)}{r.BUYUNITSIZE ? ` ${str(r.BUYUNITSIZE)}` : ''}</div></> },
                  {
                    key: 'price',
                    label: t('purch.price'),
                    align: 'right',
                    render: (r) => {
                      const p = r as BuyProduct
                      const line = lineOf(keyOf(p))
                      const base = toItem(p)
                      const price = line?.price ?? base.price
                      return base.isVariable ? (
                        <span className="qty">
                          <button type="button" onClick={() => setLine(base.key, { price: Math.max(0, price - 1000) }, base)}>−</button>
                          {fmtMoney(price)}
                          <button type="button" onClick={() => setLine(base.key, { price: price + 1000 }, base)}>+</button>
                        </span>
                      ) : (
                        fmtMoney(price)
                      )
                    },
                  },
                  {
                    key: 'qty',
                    label: t('purch.qty'),
                    align: 'right',
                    render: (r) => {
                      const p = r as BuyProduct
                      const base = toItem(p)
                      const line = lineOf(base.key)
                      const step = base.isVariable ? 0.1 : 1
                      const qty = line?.qty ?? 0
                      return (
                        <span className="qty">
                          <button type="button" disabled={qty <= 0} onClick={() => setLine(base.key, { qty: parseFloat((qty - step).toFixed(1)) }, base)}>−</button>
                          {qty}
                          <button type="button" onClick={() => setLine(base.key, { qty: parseFloat((qty + step).toFixed(1)) }, base)}>+</button>
                        </span>
                      )
                    },
                  },
                ]}
              />
            )}

            {local && (
              <form
                className="custom-item"
                onSubmit={(e) => {
                  e.preventDefault()
                  addMisc()
                }}
              >
                <h4 className="menu-cat">✏️ {t('purch.misc')}</h4>
                <div className="custom-item__row">
                  <input value={miscName} onChange={(e) => setMiscName(e.target.value)} placeholder={t('custom.name')} maxLength={50} />
                  <input type="number" inputMode="numeric" min={0} step={1000} value={miscPrice} onChange={(e) => setMiscPrice(e.target.value)} placeholder={t('custom.price')} />
                  <Button type="submit">＋ {t('custom.add')}</Button>
                </div>
              </form>
            )}
          </div>

          <aside className="purch__cart">
            <h3>
              🧾 {t('purch.cart')} <Badge tone={cart.length ? 'ok' : 'neutral'}>{cart.length}</Badge>
            </h3>
            {cart.length === 0 ? (
              <p className="muted">{t('checkout.cartEmpty')}</p>
            ) : (
              <ul className="cart">
                {cart.map((c) => (
                  <li key={c.key} className="purch__line">
                    <span className="name">
                      {c.name}
                      <div className="muted">
                        {fmtNum(c.qty)} × {fmtMoney(c.price)}
                      </div>
                      <input
                        type="date"
                        className="purch__expiry"
                        value={c.expiryDate ?? ''}
                        onChange={(e) => setLine(c.key, { expiryDate: e.target.value })}
                        title={t('purch.expiry')}
                      />
                    </span>
                    <span className="muted">{fmtMoney(c.price * c.qty)}</span>
                    <button type="button" className="soldout-btn" onClick={() => setLine(c.key, { qty: 0 })} aria-label="Remove">
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
            )}
            <div className="checkout-total">
              <span>{t('checkout.total')}</span>
              <span>{fmtMoney(total)}</span>
            </div>
            <div className="actions">
              <Button onClick={proceed} disabled={!cart.length || !account}>
                💳 {t('checkout.proceed')}
              </Button>
              {cart.length > 0 && (
                <Button variant="ghost" onClick={() => setCart([])}>
                  {t('checkout.clear')}
                </Button>
              )}
            </div>
          </aside>
        </div>
      )}
    </Panel>
  )
}

function ReceivePanel({ auth }: { auth: AuthContext }) {
  const { t, fmtNum } = useI18n()
  const [rows, setRows] = useState<OrderedPurchase[]>([])
  const [loading, setLoading] = useState(false)
  const [qty, setQty] = useState<Record<number, string>>({})
  const [msg, setMsg] = useState<Msg | null>(null)

  const load = useCallback(async () => {
    setLoading(true)
    const list = await fetchOrderedPurchases(auth)
    setLoading(false)
    setRows(list)
    setQty(Object.fromEntries(list.map((r) => [r.rowIndex, String(r.qty)])))
  }, [auth])

  useEffect(() => {
    load()
  }, [load])

  const receive = async (r: OrderedPurchase) => {
    const q = Number(qty[r.rowIndex])
    if (isNaN(q) || q < 0) return setMsg({ tone: 'error', text: t('purch.invalidQty') })
    const res = await markReceived(r.rowIndex, q, auth)
    if (res.ok) {
      setMsg({ tone: 'info', text: `${t('purch.received')} — ${r.product}` })
      load()
    } else setMsg({ tone: 'error', text: res.error ?? '' })
  }

  return (
    <Panel title={`📥 ${t('nav.purchasingReceive')}`}>
      {msg && <Notice tone={msg.tone === 'error' ? 'error' : 'info'}>{msg.text}</Notice>}
      <div className="actions">
        <Button variant="ghost" onClick={load} disabled={loading}>
          {loading ? t('common.loading') : '↻'}
        </Button>
      </div>
      <DataTable
        rows={rows as unknown as Record<string, unknown>[]}
        rowKey={(r) => String(r.rowIndex)}
        empty={t('purch.nothingToReceive')}
        columns={[
          { key: 'date', label: t('kitchen.when') },
          { key: 'product', label: t('purch.product') },
          { key: 'supplier', label: t('purch.supplier') },
          { key: 'qty', label: t('purch.ordered'), align: 'right', render: (r) => fmtNum(Number(r.qty)) },
          {
            key: 'rquantity',
            label: t('purch.receivedQty'),
            align: 'right',
            render: (r) => (
              <input
                type="number"
                inputMode="decimal"
                min={0}
                step="any"
                style={{ width: '6rem', textAlign: 'right' }}
                value={qty[Number(r.rowIndex)] ?? ''}
                onChange={(e) => setQty((q) => ({ ...q, [Number(r.rowIndex)]: e.target.value }))}
              />
            ),
          },
          { key: 'act', label: '', render: (r) => <Button onClick={() => receive(r as unknown as OrderedPurchase)}>✅ {t('purch.markReceived')}</Button> },
        ]}
      />
    </Panel>
  )
}
