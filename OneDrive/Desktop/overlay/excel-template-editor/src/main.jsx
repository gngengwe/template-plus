import React from 'react'
import ReactDOM from 'react-dom/client'
import 'handsontable/dist/handsontable.full.min.css'
import './index.css'
import App from './App'

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props)
    this.state = { error: null }
  }

  static getDerivedStateFromError(error) {
    return { error }
  }

  render() {
    if (this.state.error) {
      return (
        <div className="min-h-screen bg-slate-950 px-4 py-10 text-slate-100">
          <div className="mx-auto max-w-3xl rounded-2xl border border-rose-500/50 bg-rose-950/20 p-6">
            <h1 className="font-['Sora'] text-xl font-semibold">Application error</h1>
            <p className="mt-2 text-sm text-rose-200">The app failed to load. Please share this message.</p>
            <pre className="mt-4 overflow-auto rounded-lg bg-slate-900 p-3 text-xs text-rose-100">
              {String(this.state.error?.stack || this.state.error?.message || this.state.error)}
            </pre>
          </div>
        </div>
      )
    }

    return this.props.children
  }
}

if (typeof window !== 'undefined') {
  const mountGlobalError = (message) => {
    const root = document.getElementById('root')
    if (!root) return
    root.innerHTML = `
      <div style="min-height:100vh;background:#020617;color:#e2e8f0;padding:24px;font-family:Manrope,Segoe UI,sans-serif;">
        <div style="max-width:900px;margin:0 auto;border:1px solid rgba(244,63,94,.45);background:rgba(136,19,55,.2);border-radius:16px;padding:20px;">
          <h1 style="margin:0 0 8px;font-size:22px;">Runtime error</h1>
          <p style="margin:0 0 12px;font-size:14px;color:#fecdd3;">The app crashed before React finished rendering.</p>
          <pre style="white-space:pre-wrap;word-break:break-word;background:#0f172a;padding:12px;border-radius:10px;font-size:12px;">${String(message)}</pre>
        </div>
      </div>
    `
  }

  window.addEventListener('error', (event) => {
    if (event?.error) mountGlobalError(event.error.stack || event.error.message)
  })

  window.addEventListener('unhandledrejection', (event) => {
    const reason = event?.reason
    mountGlobalError(reason?.stack || reason?.message || reason)
  })
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </React.StrictMode>,
)