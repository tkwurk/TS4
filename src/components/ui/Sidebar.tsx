/** Left sidebar navigation. Generic over the item id string union. */
export function Sidebar<T extends string>({
  items,
  active,
  onChange,
}: {
  items: { id: T; label: string; icon?: string }[]
  active: T
  onChange: (id: T) => void
}) {
  return (
    <aside className="sidebar">
      {items.map((item) => (
        <button
          key={item.id}
          type="button"
          className={'side-item' + (item.id === active ? ' side-item--active' : '')}
          onClick={() => onChange(item.id)}
        >
          {item.icon && <span className="side-item__icon">{item.icon}</span>}
          <span className="side-item__label">{item.label}</span>
        </button>
      ))}
    </aside>
  )
}
