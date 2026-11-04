import type { ReactNode } from 'react'

export interface Column<T> {
  key: string
  label: ReactNode
  render?: (row: T) => ReactNode
  align?: 'left' | 'right' | 'center'
}

/** Generic responsive table used by report/admin screens. */
export function DataTable<T extends Record<string, unknown>>({
  columns,
  rows,
  rowKey,
  empty,
}: {
  columns: Column<T>[]
  rows: T[]
  rowKey: (row: T, index: number) => string
  empty?: ReactNode
}) {
  if (!rows.length) return <p className="muted">{empty ?? '—'}</p>
  return (
    <div className="table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c.key} style={{ textAlign: c.align ?? 'left' }}>
                {c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={rowKey(r, i)}>
              {columns.map((c) => (
                <td key={c.key} style={{ textAlign: c.align ?? 'left' }}>
                  {c.render ? c.render(r) : String(r[c.key] ?? '')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
