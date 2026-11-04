/**
 * Checkout data access — submits through the Seam 1 API only (api.checkout),
 * which routes to the Orchestrator's checkout.completed pipeline. Never touches
 * Sheets directly (see /AGENTS.md).
 */
import { api } from '@/services/api'
import type { ApiEnvelope, AuthContext } from '@/services/api'
import type { CartItem, CheckoutResult } from './types'

export function submitCheckout(
  items: CartItem[],
  payment: string,
  auth: AuthContext,
): Promise<ApiEnvelope<CheckoutResult>> {
  const orderData = {
    items: items.map((i) => ({ name: i.name, qty: i.qty, price: i.price })),
    payment,
    total: items.reduce((sum, i) => sum + i.price * i.qty, 0),
  }
  return api.checkout<CheckoutResult>({ orderData }, auth)
}
