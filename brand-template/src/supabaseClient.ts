import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.REACT_APP_SUPABASE_URL!
const supabaseAnonKey = process.env.REACT_APP_SUPABASE_ANON_KEY!

// In production, route all Supabase traffic through Vercel's proxy (/api/sb/*)
// so the browser never makes direct requests to supabase.co
const effectiveUrl =
  typeof window !== 'undefined' && window.location.hostname !== 'localhost'
    ? `${window.location.origin}/api/sb`
    : supabaseUrl

export const supabase = createClient(effectiveUrl, supabaseAnonKey)
