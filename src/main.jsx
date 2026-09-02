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
// Picks the ELEMENT rather than the component type on purpose: a capitalised
// `const RootComponent = ...` reads to react-refresh as a component declared
// in a file with no exports, which it (correctly, in general) flags as
// un-hot-reloadable. This is the app's entry point, so it genuinely has
// nothing to export - choosing the element sidesteps that without suppressing
// the rule anywhere. Behaviour is identical.
const params = new URLSearchParams(location.search)
const rootElement =
  params.get('testpattern') === '1' ? <TestPatternView /> :
  params.get('quicklayout') === '1' ? <QuickLayoutView /> :
  <App />

createRoot(document.getElementById('root')).render(
  <StrictMode>
    {rootElement}
  </StrictMode>,
)
