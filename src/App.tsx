import { useMemo, useState } from 'react'
import { useAuth, LoginPanel } from '@/features/auth'
import { CheckoutPanel } from '@/features/checkout'
import { OrdersPanel } from '@/features/orders'
import { ReportsPanel } from '@/features/reports'
import { ComingSoon, visibleNavItems } from '@/features/navigation'
import { Button, Sidebar } from '@/components/ui'
import { useI18n } from '@/i18n'
import './App.css'

function LanguageToggle() {
  const { lang, setLang } = useI18n()
  return (
    <span className="lang-toggle">
      <button
        type="button"
        className={'lang-opt' + (lang === 'id' ? ' lang-opt--active' : '')}
        onClick={() => setLang('id')}
      >
        ID
      </button>
      <button
        type="button"
        className={'lang-opt' + (lang === 'en' ? ' lang-opt--active' : '')}
        onClick={() => setLang('en')}
      >
        EN
      </button>
    </span>
  )
}

function App() {
  const { session, loading, error, login, logout } = useAuth()
  const { t } = useI18n()
  const [active, setActive] = useState('orders')

  const navItems = useMemo(() => (session ? visibleNavItems(session) : []), [session])

  if (!session) {
    return (
      <main className="app">
        <header className="app__header">
          <h1>TS4</h1>
          <p className="subtitle">{t('app.subtitle')}</p>
          <LanguageToggle />
        </header>
        <LoginPanel onLogin={login} loading={loading} error={error} />
      </main>
    )
  }

  const auth = { username: session.username, deviceId: session.deviceId, outlet: session.outlet }
  const currentId = navItems.some((i) => i.id === active) ? active : navItems[0]?.id
  const current = navItems.find((i) => i.id === currentId)

  const renderScreen = () => {
    switch (currentId) {
      case 'orders':
        return <CheckoutPanel session={session} />
      case 'adminOrders':
        return <OrdersPanel auth={auth} />
      case 'monthlySalesReport':
        return <ReportsPanel auth={auth} />
      default:
        return <ComingSoon title={current ? t(current.labelKey) : ''} />
    }
  }

  return (
    <div className="shell">
      <header className="topbar">
        <h1 className="topbar__logo">TS4</h1>
        <div className="actions">
          <span className="status">
            {session.username}
            {session.outlet ? ` · ${session.outlet}` : ''} · {session.role}
          </span>
          <LanguageToggle />
          <Button variant="ghost" onClick={logout}>
            {t('app.signOut')}
          </Button>
        </div>
      </header>

      <div className="shell__body">
        <Sidebar
          items={navItems.map((i) => ({ id: i.id, label: t(i.labelKey), icon: i.icon }))}
          active={currentId ?? ''}
          onChange={setActive}
        />
        <section className="content">{renderScreen()}</section>
      </div>
    </div>
  )
}

export default App
