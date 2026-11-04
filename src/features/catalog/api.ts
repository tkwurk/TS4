/**
 * Catalog data access — talks to the backend only through the Seam 1 API client.
 * Never calls the Google Sheets API directly (see /AGENTS.md).
 */
import { api } from '@/services/api'
import type { ApiEnvelope, AuthContext } from '@/services/api'
import type { CatalogPipeline, CatalogResult } from './types'

/**
 * Fetch active products for an outlet (empty outlet = all). The catalog.requested
 * event runs through the Orchestrator, so the product list is nested under
 * steps.catalog.data — this unwraps it into a flat CatalogResult envelope.
 */
export async function fetchProducts(
  opts: { outlet?: string; type?: string } = {},
  auth?: AuthContext,
): Promise<ApiEnvelope<CatalogResult>> {
  const res = await api.products<CatalogPipeline>(opts, auth)
  if (!res.success) {
    return { success: false, data: null, error: res.error, code: res.code }
  }
  const data = res.data?.steps?.catalog?.data ?? { outlet: null, count: 0, products: [] }
  return { success: true, data, error: null, code: res.code }
}
