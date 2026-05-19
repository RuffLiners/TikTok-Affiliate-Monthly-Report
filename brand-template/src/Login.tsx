import React, { useState, useEffect } from 'react'
import { supabase } from './supabaseClient'

export default function Login() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  useEffect(() => {
    supabase.auth.signOut({ scope: 'local' }).catch(() => {})
  }, [])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setLoading(true)
    setError('')
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) setError(error.message)
    setLoading(false)
  }

  return (
    <div style={{ minHeight: '100vh', background: '#0d1117', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: '#1a2035', border: '1px solid #2d3748', borderRadius: 12, padding: 40, width: 360 }}>
        <div style={{ textAlign: 'center', marginBottom: 32 }}>
          <div style={{ fontSize: 28 }}>📦</div>
          <h1 style={{ color: '#f9fafb', fontSize: 20, fontWeight: 700, margin: '8px 0 4px' }}>Your Brand</h1>
          <p style={{ color: '#6b7280', fontSize: 13, margin: 0 }}>Creator Hub · TikTok Shop</p>
        </div>
        <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <div>
            <label style={{ fontSize: 12, fontWeight: 600, color: '#9ca3af', display: 'block', marginBottom: 6 }}>EMAIL</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="you@yourbrand.com" required
              style={{ width: '100%', background: '#0d1117', border: '1px solid #2d3748', borderRadius: 6, color: '#e5e7eb', padding: '10px 12px', fontSize: 13, boxSizing: 'border-box', outline: 'none' }} />
          </div>
          <div>
            <label style={{ fontSize: 12, fontWeight: 600, color: '#9ca3af', display: 'block', marginBottom: 6 }}>PASSWORD</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" required
              style={{ width: '100%', background: '#0d1117', border: '1px solid #2d3748', borderRadius: 6, color: '#e5e7eb', padding: '10px 12px', fontSize: 13, boxSizing: 'border-box', outline: 'none' }} />
          </div>
          {error && (
            <div style={{ background: '#ef444422', border: '1px solid #ef4444', borderRadius: 6, padding: '8px 12px', fontSize: 12, color: '#ef4444' }}>{error}</div>
          )}
          <button type="submit" disabled={loading}
            style={{ background: loading ? '#1e3a5f' : '#3b82f6', color: '#fff', border: 'none', borderRadius: 7, padding: '11px 0', fontSize: 14, fontWeight: 700, cursor: loading ? 'not-allowed' : 'pointer', marginTop: 4 }}>
            {loading ? 'Please wait…' : 'Sign In'}
          </button>
        </form>
        <p style={{ textAlign: 'center', fontSize: 12, color: '#4b5563', marginTop: 24, marginBottom: 0 }}>Contact your admin to get access.</p>
      </div>
    </div>
  )
}
