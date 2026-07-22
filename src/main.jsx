import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import TestPatternView from './testPattern/TestPatternView.tsx'
import QuickLayoutView from './quickLayout/QuickLayoutView.tsx'

// No router in this single-page app: the animated test pattern and the Quick
// Panel Layout calculator each open in their own tab via window.open(...) with
// the query string below, independent of the main editor. Detect that here
// and render the dedicated view instead of the normal editor.
const params = new URLSearchParams(location.search)
const RootComponent =
  params.get('testpattern') === '1' ? TestPatternView :
  params.get('quicklayout') === '1' ? QuickLayoutView :
  App

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <RootComponent />
  </StrictMode>,
)
