import { api } from '@/services/api'
import type { AuthContext } from '@/services/api'

export interface BuyProduct {
  productCode: string
  name: string
  category?: string
  navCategory?: string
  unit?: string
  buyPrice?: number
  supplier?: string
}

export interface PrefillItem {
  productCode: string
  name?: string
  unit?: string
  quantity: number
  [k: string]: unknown
}

export interface SummaryItem {
  productCode?: string
  name?: string
  unit?: string
  quantity?: number
  [k: string]: unknown
}

export interface KeyfigureItem {
  name: string
  productCode: string
  category?: string
  navCategory?: string
  unit?: string
  totalPurchased: number
  totalConsumed: number
  totalSold: number
  remaining: number
  reorderQty?: number
  supplier?: string
  hasRecipe?: boolean
}

export interface DailyClosingBundle {
  products: BuyProduct[]
  autoPrefill: { date: string; totalItems: number; items: PrefillItem[]; missingProductCodes?: string[] }
  summary: { totalItems: number; totalQuantity: number; items: SummaryItem[] }
  keyfigures?: { items?: KeyfigureItem[]; stockSource?: string; historyAsOfDate?: string | null; needsBootstrap?: boolean }
  smartPicks?: { orders?: unknown[]; consumption?: unknown[] }
}

export interface ConsumptionHistoryRow extends Record<string, unknown> {
  date: string
  productCode: string
  name: string
  unit: string
  quantity: number
}

export const getDailyClosingBundle = (auth: AuthContext) =>
  api.consumptionClosing<DailyClosingBundle>(auth)

export const recordConsumption = (items: { productCode: string; quantity: number }[], auth: AuthContext) =>
  api.recordConsumption<{ message?: string }>(
    { items, clientRequestId: `ic-${Date.now()}-${Math.random().toString(36).slice(2, 10)}` },
    auth,
  )

export const getConsumptionHistory = (startDate: string, endDate: string, auth: AuthContext) =>
  api.consumptionHistory<{ rows: ConsumptionHistoryRow[] }>(startDate, endDate, auth)
