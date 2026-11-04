import { callPositional, callSecure, rpc } from '@/services/api'
import type { AuthContext } from '@/services/api'

const parseMaybeJson = <T,>(d: unknown): T | null => {
  if (d == null) return null
  if (typeof d === 'string') {
    try {
      return JSON.parse(d) as T
    } catch {
      return null
    }
  }
  return d as T
}

// ── P&L ──
export interface PnlRow {
  label?: string
  date?: string
  sales: number
  purchases: number
  expenses: number
  profit: number
  isFuture?: boolean
  isPartial?: boolean
}
export interface PnlCategory {
  category: string
  sales: number
  purchases: number
  expenses: number
  salesPct?: number
  profit: number
}
export interface PnlBundle {
  year: number
  daily: PnlRow[]
  monthly: PnlRow[]
  category: { categories: PnlCategory[]; totalExpenses: number }
  monthlyCategory?: Record<string, unknown>
}

export async function fetchPnl(year: number, auth: AuthContext): Promise<{ data: PnlBundle | null; error: string | null }> {
  const r = await callSecure<PnlBundle>('getPnLBundleSecure', { year }, auth)
  if (!r.ok) return { data: null, error: r.error }
  // apiOk copies the bundle's keys onto the envelope; accept either shape.
  const d = (r.data as unknown as { data?: PnlBundle }) ?? {}
  return { data: (d.data as PnlBundle) ?? (r.data as PnlBundle), error: null }
}

// ── Usage ──
export interface UsageRow {
  month: string
  category: string
  name: string
  qty: number
}
export async function fetchUsage(auth: AuthContext): Promise<UsageRow[]> {
  const r = await callPositional<unknown>('getMonthlyUsageReport', [auth.deviceId, auth.username])
  return Array.isArray(r.data) ? (r.data as UsageRow[]) : []
}

// ── Sales by date ──
export interface SalesSummary {
  date: string
  summary: Record<string, { qty?: number; total?: number; category?: string; [k: string]: unknown }>
  totalSales: number
  totalSalesByCash: number
  totalSalesByNonCash: number
  salesByPaymentMethod: Record<string, number>
  ordersCount: number
  itemsCount?: number
  totalDeposits: number
  totalDepositsByQris?: number
  missedOrdersCount: number
  missedOrdersTotal: number
}
export interface OrderLine {
  orderNumber?: string
  time?: string
  product?: string
  qty?: number
  price?: number
  total?: number
  category?: string
  payment?: string
  username?: string
  [k: string]: unknown
}

export async function fetchSalesSummary(date: string, auth: AuthContext): Promise<SalesSummary | null> {
  const r = await rpc<unknown>('getSalesSummaryByDate', [{ dateString: date, _authContext: auth }])
  return r.success ? parseMaybeJson<SalesSummary>(r.data) : null
}

export async function fetchOrderLines(date: string, auth: AuthContext): Promise<OrderLine[]> {
  const r = await rpc<unknown>('getOrderLinesByDate', [{ dateString: date, _authContext: auth }])
  const d = r.success ? parseMaybeJson<{ lines?: OrderLine[] }>(r.data) : null
  return d?.lines ?? []
}
