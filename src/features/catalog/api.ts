/**
 * Catalog data access — talks to the backend only through the Seam 1 API client.
 * Never calls the Google Sheets API directly (see /AGENTS.md).
 */
import { api } from '@/services/api'
import type { ApiEnvelope, AuthContext } from '@/services/api'
import type { CatalogResult } from './types'

/** Fetch active products for an outlet (empty outlet = all). */
export function fetchProducts(
  opts: { outlet?: string; type?: string } = {},
  auth?: AuthContext,
): Promise<ApiEnvelope<CatalogResult>> {
  return api.products<CatalogResult>(opts, auth)
}
