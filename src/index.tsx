import React from 'react'
import ReactDOM from 'react-dom/client'
import './index.css'
import AuthWrapper from './AuthWrapper'
import TikTokShopReporter from './TikTokShopReporter'

const root = ReactDOM.createRoot(
  document.getElementById('root') as HTMLElement
)

root.render(
  <React.StrictMode>
    <AuthWrapper>
      <TikTokShopReporter />
    </AuthWrapper>
  </React.StrictMode>
)
