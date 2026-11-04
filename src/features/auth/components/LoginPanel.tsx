import { useState, type FormEvent } from 'react'
import { apiConfig } from '@/services/api'
import { Button, Notice, Panel } from '@/components/ui'

/** Operator sign-in form. Submits credentials via the auth hook's `onLogin`. */
export function LoginPanel({
  onLogin,
  loading,
  error,
}: {
  onLogin: (username: string, password: string) => void
  loading: boolean
  error: string | null
}) {
  const configured = Boolean(apiConfig.appsScriptUrl)
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')

  const submit = (e: FormEvent) => {
    e.preventDefault()
    onLogin(username, password)
  }

  return (
    <Panel title="Sign in">
      {!configured && (
        <Notice tone="warn">
          Set <code>VITE_APPS_SCRIPT_URL</code> in <code>.env.local</code> to sign in.
        </Notice>
      )}
      <form onSubmit={submit}>
        <label className="ui-field">
          Username
          <input value={username} onChange={(e) => setUsername(e.target.value)} autoComplete="username" />
        </label>
        <label className="ui-field">
          Password
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoComplete="current-password"
          />
        </label>
        <div className="actions">
          <Button type="submit" disabled={!configured || loading || !username || !password}>
            {loading ? 'Signing in…' : 'Sign in'}
          </Button>
        </div>
      </form>
      {error && <Notice tone="error">{error}</Notice>}
    </Panel>
  )
}
