import { productName, type ProductRow } from '@/features/catalog'

export interface CartItem {
  key: string
  name: string
  price: number
  qty: number
}

/** Shape returned by the checkout.completed pipeline (apiOk data). */
export interface CheckoutResult {
  eventType?: string
  requestId?: string
  steps?: {
    sales?: {
      orderNumber?: string
      data?: { orderNumber?: string; orderNumbers?: string[] }
    }
  }
}

const PRICE_KEYS = ['PRICE', 'Price', 'price', 'UnitPrice', 'Unit Price']

/** Best-effort numeric price from a header-keyed product row. */
export function productPrice(row: ProductRow): number {
  for (const k of PRICE_KEYS) {
    if (k in row) {
      const n = Number(row[k])
      if (!Number.isNaN(n)) return n
    }
  }
  return 0
}

/** Best-effort product image URL from a header-keyed row (Drive file id or URL). */
export function productImage(row: ProductRow): string | null {
  const key = Object.keys(row).find((k) => {
    const u = k.toUpperCase()
    return u === 'IMAGEFILEID' || u === 'IMAGE' || u === 'IMAGEURL'
  })
  const raw = key ? row[key] : undefined
  const id = raw != null ? String(raw).trim() : ''
  if (!id) return null
  if (/^(https?:|data:)/i.test(id)) return id
  return `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w200`
}

/** Best-effort supplier from a header-keyed product row. */
export function productSupplier(row: ProductRow): string {
  const key = Object.keys(row).find((k) => k.toUpperCase() === 'SUPPLIER')
  const value = key ? row[key] : undefined
  return value != null && String(value).trim() !== '' ? String(value).trim() : 'Other'
}

/** Best-effort category from a header-keyed product row. */
export function productCategory(row: ProductRow): string {
  const key = Object.keys(row).find((k) => /^category$|^cat$|nav ?category/i.test(k))
  const value = key ? row[key] : undefined
  return value != null && String(value).trim() !== '' ? String(value).trim() : 'Other'
}

/** Best-effort product type ('Sell' / 'Buy' / …); defaults to 'Sell'. */
export function productType(row: ProductRow): string {
  const key = Object.keys(row).find((k) => k.toUpperCase() === 'TYPE')
  const value = key ? row[key] : undefined
  return value != null && String(value).trim() !== '' ? String(value).trim() : 'Sell'
}

/** Stable identity for a product row — product code if present, else the name. */
export function productKey(row: ProductRow): string {
  const codeKey = Object.keys(row).find((k) => /product ?code|^code$/i.test(k))
  const code = codeKey ? row[codeKey] : undefined
  return code != null && code !== '' ? String(code) : productName(row)
}

/** Pull the order number out of the nested pipeline result, if present. */
export function orderNumberOf(result: CheckoutResult | null | undefined): string | undefined {
  const sales = result?.steps?.sales
  return sales?.data?.orderNumber ?? sales?.orderNumber
}
