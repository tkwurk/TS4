/**
 * App navigation model — mirrors the original ui_navigation.html tab set and
 * permission names (allowedTabs). `id` is the permission key used for role
 * gating; `ready` marks screens already migrated to the PWA.
 */
export interface NavItem {
  id: string
  labelKey: string
  icon: string
  ready?: boolean
}

export const NAV_ITEMS: NavItem[] = [
  { id: 'orders', labelKey: 'nav.orders', icon: '🧾', ready: true },
  { id: 'adminOrders', labelKey: 'nav.adminOrders', icon: '📋', ready: true },
  { id: 'monthlySalesReport', labelKey: 'nav.monthlySalesReport', icon: '📅', ready: true },
  { id: 'kitchenCooking', labelKey: 'nav.kitchenCooking', icon: '👨‍🍳' },
  { id: 'purchasing', labelKey: 'nav.purchasing', icon: '📦' },
  { id: 'purchasingReceive', labelKey: 'nav.purchasingReceive', icon: '📥' },
  { id: 'purchasingLocal', labelKey: 'nav.purchasingLocal', icon: '🏪' },
  { id: 'stockReport', labelKey: 'nav.stockReport', icon: '📊' },
  { id: 'inventoryConsumption', labelKey: 'nav.inventoryConsumption', icon: '🧮' },
  { id: 'deposits', labelKey: 'nav.deposits', icon: '💰' },
  { id: 'investments', labelKey: 'nav.investments', icon: '💎' },
  { id: 'profitLoss', labelKey: 'nav.profitLoss', icon: '📈' },
  { id: 'usageReport', labelKey: 'nav.usageReport', icon: '📉' },
  { id: 'attendance', labelKey: 'nav.attendance', icon: '🕒' },
  { id: 'checklist', labelKey: 'nav.checklist', icon: '✅' },
]
