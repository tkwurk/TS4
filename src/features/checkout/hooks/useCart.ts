import { useCallback, useMemo, useState } from 'react'
import type { CartItem } from '../types'

/** Isolated cart state + operations (kept out of the view component). */
export function useCart() {
  const [items, setItems] = useState<CartItem[]>([])

  const add = useCallback((item: Omit<CartItem, 'qty'>) => {
    setItems((prev) => {
      const i = prev.findIndex((p) => p.key === item.key)
      if (i >= 0) {
        const next = [...prev]
        next[i] = { ...next[i], qty: next[i].qty + 1 }
        return next
      }
      return [...prev, { ...item, qty: 1 }]
    })
  }, [])

  const setQty = useCallback((key: string, qty: number) => {
    setItems((prev) => prev.map((p) => (p.key === key ? { ...p, qty } : p)).filter((p) => p.qty > 0))
  }, [])

  const remove = useCallback((key: string) => {
    setItems((prev) => prev.filter((p) => p.key !== key))
  }, [])

  const clear = useCallback(() => setItems([]), [])

  const total = useMemo(() => items.reduce((sum, i) => sum + i.price * i.qty, 0), [items])
  const count = useMemo(() => items.reduce((sum, i) => sum + i.qty, 0), [items])
  const qtyOf = useCallback((key: string) => items.find((p) => p.key === key)?.qty ?? 0, [items])

  return { items, add, setQty, remove, clear, total, count, qtyOf }
}
