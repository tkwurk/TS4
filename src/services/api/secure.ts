/**
 * Convenience wrapper over the RPC bridge for the common "*Secure(params)" shape:
 * one params object carrying `_authContext`. Unwraps the server's own
 * `{ success, ...payload }` object (sanitizeForClient) or apiOk envelope.
 */
import { rpc } from './rpc'
import type { AuthContext } from './types'

export interface SecureResult<T> {
  ok: boolean
  data: T | null
  error: string | null
}

export async function callSecure<T = Record<string, unknown>>(
  fn: string,
  params: Record<string, unknown>,
  auth: AuthContext,
): Promise<SecureResult<T>> {
  const res = await rpc<unknown>(fn, [{ ...params, _authContext: auth }])
  if (!res.success) return { ok: false, data: null, error: res.error ?? 'Request failed' }
  // Server functions return either their own object (possibly with success:false)
  // or a plain value that the bridge wrapped in apiOk.
  const d = res.data as Record<string, unknown> | null
  if (d && typeof d === 'object' && 'success' in d && d.success === false) {
    return { ok: false, data: null, error: String(d.error ?? d.message ?? 'Request failed') }
  }
  return { ok: true, data: (d as T) ?? null, error: null }
}

/** For legacy functions that take positional args instead of a params object. */
export async function callPositional<T = unknown>(fn: string, args: unknown[]): Promise<SecureResult<T>> {
  const res = await rpc<T>(fn, args)
  if (!res.success) return { ok: false, data: null, error: res.error ?? 'Request failed' }
  return { ok: true, data: res.data, error: null }
}
