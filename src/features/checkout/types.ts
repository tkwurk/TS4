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
