import { useAuth, LoginPanel } from '@/features/auth'
import { CheckoutPanel } from '@/features/checkout'
import { CatalogPanel } from '@/features/catalog'
import { ReportsPanel } from '@/features/reports'
import { Button } from '@/components/ui'
import './App.css'

function App() {
  const { session, loading, error, login, logout } = useAuth()

  return (
    <main className="app">
      <header className="app__header">
        <h1>TS4</h1>
        <p className="subtitle">Point of sale · Vite + React · Google Sheets backend</p>
      </header>

      {session ? (
        <>
          <div className="actions">
            <span className="status">
              Signed in as <strong>{session.username}</strong>
              {session.outlet ? ` · ${session.outlet}` : ''} · {session.role}
            </span>
            <Button variant="ghost" onClick={logout}>
              Sign out
            </Button>
          </div>

          <CheckoutPanel session={session} />
          <ReportsPanel auth={{ username: session.username, deviceId: session.deviceId, outlet: session.outlet }} />
          <CatalogPanel />
        </>
      ) : (
        <LoginPanel onLogin={login} loading={loading} error={error} />
      )}
    </main>
  )
}

export default App
