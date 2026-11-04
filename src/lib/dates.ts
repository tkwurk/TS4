/**
 * Local-time date helpers. `Date#toISOString()` is UTC, so using it for
 * date-only defaults shifts the day after local midnight — always use these
 * for YYYY-MM-DD values shown in forms or sent as business dates.
 */
const pad = (n: number) => String(n).padStart(2, '0')

/** YYYY-MM-DD in the device's local timezone. */
export function localDate(d: Date = new Date()): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
}

/** YYYY-MM-DDTHH:mm in local time, for <input type="datetime-local">. */
export function localDateTime(d: Date = new Date()): string {
  return `${localDate(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`
}

/** Shift a YYYY-MM-DD string by N days (local calendar). */
export function addDays(date: string, days: number): string {
  const [y, m, dd] = date.split('-').map(Number)
  const d = new Date(y, (m || 1) - 1, dd || 1)
  d.setDate(d.getDate() + days)
  return localDate(d)
}
