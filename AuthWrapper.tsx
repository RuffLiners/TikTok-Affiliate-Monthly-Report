import React, { useEffect, useState } from 'react'
import { supabase } from './supabaseClient'

function SetNewPassword({ onDone }: { onDone?: () => void }) {
  const [password, setPassword] = useState('')
  const [confirm, setConfirm] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState(false)

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (password !== confirm) { setError('Passwords do not match'); return }
    if (password.length < 6) { setError('Password must be at least 6 characters'); return }
    setLoading(true); setError('')
    const { error } = await supabase.auth.updateUser({ password })
    if (error) { setError(error.message); setLoading(false); return }
    setDone(true)
    setLoading(false)
    setTimeout(() => onDone?.(), 1500)
  }

  return (
    <div style={{ minHeight: '100vh', background: '#0d1117', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: '#1a2035', border: '1px solid #2d3748', borderRadius: 12, padding: 40, width: 360 }}>
        <div style={{ textAlign: 'center', marginBottom: 28 }}>
          <div style={{ fontSize: 28 }}>🔐</div>
          <h1 style={{ color: '#f9fafb', fontSize: 18, fontWeight: 700, margin: '8px 0 4px' }}>Set New Password</h1>
          <p style={{ color: '#6b7280', fontSize: 13, margin: 0 }}>Choose a new password for your account.</p>
        </div>
        {done ? (
          <div style={{ background: '#10b98122', border: '1px solid #10b981', borderRadius: 8, padding: '14px 16px', color: '#10b981', fontSize: 13, textAlign: 'center' }}>✅ Password updated! You can now sign in.</div>
        ) : (
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div>
              <label style={{ fontSize: 12, fontWeight: 600, color: '#9ca3af', display: 'block', marginBottom: 6 }}>NEW PASSWORD</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" required
                style={{ width: '100%', background: '#0d1117', border: '1px solid #2d3748', borderRadius: 6, color: '#e5e7eb', padding: '10px 12px', fontSize: 13, boxSizing: 'border-box', outline: 'none' }} />
            </div>
            <div>
              <label style={{ fontSize: 12, fontWeight: 600, color: '#9ca3af', display: 'block', marginBottom: 6 }}>CONFIRM PASSWORD</label>
              <input type="password" value={confirm} onChange={e => setConfirm(e.target.value)} placeholder="••••••••" required
                style={{ width: '100%', background: '#0d1117', border: '1px solid #2d3748', borderRadius: 6, color: '#e5e7eb', padding: '10px 12px', fontSize: 13, boxSizing: 'border-box', outline: 'none' }} />
            </div>
            {error && <div style={{ background: '#ef444422', border: '1px solid #ef4444', borderRadius: 6, padding: '8px 12px', fontSize: 12, color: '#ef4444' }}>{error}</div>}
            <button type="submit" disabled={loading}
              style={{ background: loading ? '#1e3a5f' : '#3b82f6', color: '#fff', border: 'none', borderRadius: 7, padding: '11px 0', fontSize: 14, fontWeight: 700, cursor: loading ? 'not-allowed' : 'pointer', marginTop: 4 }}>
              {loading ? 'Saving…' : 'Set Password'}
            </button>
          </form>
        )}
      </div>
    </div>
  )
}

// Always renders children — no login gate.
// Recovery/invite links still work. Admin controls are shown inside the app based on session.
export default function AuthWrapper({ children }: { children: React.ReactNode }) {
  const [recoveryMode, setRecoveryMode] = useState(false)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    const hashParams = new URLSearchParams(window.location.hash.slice(1))
    const searchParams = new URLSearchParams(window.location.search)
    const urlType = hashParams.get('type') || searchParams.get('type')
    if (urlType === 'invite' || urlType === 'recovery') setRecoveryMode(true)

    supabase.auth.getSession()
      .then(() => setLoading(false))
      .catch(() => setLoading(false))

    const { data: { subscription } } = supabase.auth.onAuthStateChange((event) => {
      if (event === 'PASSWORD_RECOVERY') setRecoveryMode(true)
      else setRecoveryMode(false)
    })

    return () => subscription.unsubscribe()
  }, [])

  if (loading) {
    return (
      <div style={{ minHeight: '100vh', background: '#0d1117', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 12 }}>
        <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
        <div style={{ width: 28, height: 28, border: '3px solid #1e2433', borderTopColor: '#3b82f6', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} />
        <div style={{ color: '#6b7280', fontSize: 13 }}>Loading…</div>
      </div>
    )
  }

  if (recoveryMode) return <SetNewPassword onDone={() => setRecoveryMode(false)} />
  return <>{children}</>
}
