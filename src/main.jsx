import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import TestPatternView from './testPattern/TestPatternView.tsx'

// No router in this single-page app: the animated test pattern opens in its
// own tab via window.open(...?testpattern=1) with the project handed off
// through localStorage (see App.tsx openAnimatedTestPatternTab). Detect that
// here and render the dedicated view instead of the normal editor.
const params = new URLSearchParams(location.search)
const RootComponent = params.get('testpattern') === '1' ? TestPatternView : App

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <RootComponent />
  </StrictMode>,
)
