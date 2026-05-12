'use client'

import { useState, useRef, useMemo } from 'react'
import * as XLSX from 'xlsx'

const HIDDEN_COLUMNS = new Set(['Open origin', 'Close origin', 'SL', 'TP', 'Margin', 'Comment'])

type TradeRow = Record<string, any>

export default function Home() {
  const [fileName, setFileName] = useState('')
  const [rows, setRows] = useState<TradeRow[]>([])
  const [error, setError] = useState('')
  const [page, setPage] = useState(0)
  const [loading, setLoading] = useState(false)
  const [sortKey, setSortKey] = useState<string>('')
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc')
  const fileInputRef = useRef<HTMLInputElement>(null)

  function triggerFileUpload() { fileInputRef.current?.click() }

  function formatExcelDate(value: any) {
    if (typeof value !== 'number') return value
    const date = XLSX.SSF.parse_date_code(value)
    if (!date) return value
    const jsDate = new Date(date.y, date.m - 1, date.d, date.H, date.M, date.S)
    return `${String(jsDate.getDate()).padStart(2,'0')}.${String(jsDate.getMonth()+1).padStart(2,'0')}.${jsDate.getFullYear()} ${String(jsDate.getHours()).padStart(2,'0')}:${String(jsDate.getMinutes()).padStart(2,'0')}:${String(jsDate.getSeconds()).padStart(2,'0')}`
  }

  function excelDateToApiDate(value: any) {
    if (typeof value !== 'number') return ''
    const date = XLSX.SSF.parse_date_code(value)
    if (!date) return ''
    return `${date.y}-${String(date.m).padStart(2,'0')}-${String(date.d).padStart(2,'0')}`
  }

  function getTaxExemption(openValue: any, closeValue: any) {
    if (typeof openValue !== 'number' || typeof closeValue !== 'number') return ''
    const open = XLSX.SSF.parse_date_code(openValue)
    const close = XLSX.SSF.parse_date_code(closeValue)
    if (!open || !close) return ''
    const openDate = new Date(open.y, open.m - 1, open.d)
    const closeDate = new Date(close.y, close.m - 1, close.d)
    const yearDiff = closeDate.getFullYear() - openDate.getFullYear()
    if (yearDiff > 3) return 'Yes'
    if (yearDiff < 3) return 'No'
    const anniversary = new Date(closeDate.getFullYear(), openDate.getMonth(), openDate.getDate())
    return closeDate > anniversary ? 'Yes' : 'No'
  }

  const fxCache = useRef(new Map<string, Promise<number | ''>>())
  async function fetchFx(date: string): Promise<number | ''> {
    if (!date) return ''
    const cached = fxCache.current.get(date)
    if (cached) return cached
    const promise = (async () => {
      try {
        const res = await fetch(`/api/fx?date=${date}`)
        const text = await res.text()
        let rate: number | '' = ''
        try { rate = JSON.parse(text).rate ?? '' } catch { rate = '' }
        return rate !== '' ? Number(rate) : ''
      } catch { return '' }
    })()
    fxCache.current.set(date, promise)
    return promise
  }

  function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    setLoading(true); setFileName(file.name); setRows([]); setError(''); setPage(0)
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheetName = workbook.SheetNames.find(name => name.toLowerCase().includes('closed'))
        if (!sheetName) { setError('No CLOSED POSITION sheet found'); setLoading(false); return }
        const sheet = workbook.Sheets[sheetName]
        const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][]
        const KNOWN_HEADERS = new Set(['Symbol', 'Open time', 'Close time', 'Volume'])
        const headerIndex = raw.findIndex(row => Array.isArray(row) && row.filter(cell => KNOWN_HEADERS.has(String(cell).trim())).length >= 3)
        if (headerIndex === -1) { setError('Could not find header row — unexpected file format'); setLoading(false); return }
        const headers = raw[headerIndex] as string[]
        const REQUIRED_COLUMNS = ['Symbol', 'Open time', 'Close time', 'Purchase value', 'Sale value']
        const missingColumns = REQUIRED_COLUMNS.filter(col => !headers.includes(col))
        if (missingColumns.length > 0) { setError(`Missing expected columns: ${missingColumns.join(', ')}`); setLoading(false); return }
        const dataRows = raw.slice(headerIndex + 1)
        const cleaned = dataRows.filter(r => { const first = String(r?.[0] ?? '').toLowerCase().trim(); return first && first !== 'total' })
        const json = await Promise.all(cleaned.map(async row => {
          const obj: TradeRow = {}
          headers.forEach((h, i) => { obj[h] = row[i] })
          obj['Tax exemption'] = getTaxExemption(obj['Open time'], obj['Close time'])
          const openDate = excelDateToApiDate(obj['Open time'])
          const closeDate = excelDateToApiDate(obj['Close time'])
          const [openFx, closeFx] = await Promise.all([fetchFx(openDate), fetchFx(closeDate)])
          obj['Open date fx'] = openFx ? openFx.toFixed(3) : ''
          obj['Close date fx'] = closeFx ? closeFx.toFixed(3) : ''
          const purchase = Number(obj['Purchase value'] ?? 0)
          const sale = Number(obj['Sale value'] ?? 0)
          obj['Purchase in CZK'] = openFx ? (purchase * openFx).toFixed(2) : ''
          obj['Sale in CZK'] = closeFx ? (sale * closeFx).toFixed(2) : ''
          return obj
        }))
        setSortKey('')
        setSortDir('asc')
        setRows(json)
      } catch (err) { console.error(err); setError('Failed to process file') }
      finally { setLoading(false) }
    }
    reader.readAsArrayBuffer(file)
  }

  function exportToExcel() {
    if (!rows.length) return
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Trades')
    XLSX.writeFile(wb, 'xtb_trades_export.xlsx')
  }

  const visibleKeys = useMemo(() => rows.length > 0 ? Object.keys(rows[0]).filter(k => !HIDDEN_COLUMNS.has(k)) : [], [rows])

  const sortedRows = useMemo(() => {
    if (!sortKey) return rows
    return [...rows].sort((a, b) => {
      const av = a[sortKey], bv = b[sortKey]
      const aNum = Number(av), bNum = Number(bv)
      const bothNumbers = !isNaN(aNum) && !isNaN(bNum)
      const result = bothNumbers ? aNum - bNum : String(av ?? '').localeCompare(String(bv ?? ''))
      return sortDir === 'asc' ? result : -result
    })
  }, [rows, sortKey, sortDir])

  const totalPurchaseCZK   = rows.reduce((s, r) => s + Number(r['Purchase in CZK'] ?? 0), 0)
  const exemptPurchaseCZK  = rows.reduce((s, r) => r['Tax exemption'] === 'Yes' ? s + Number(r['Purchase in CZK'] ?? 0) : s, 0)
  const taxablePurchaseCZK = rows.reduce((s, r) => r['Tax exemption'] === 'No'  ? s + Number(r['Purchase in CZK'] ?? 0) : s, 0)
  const totalSaleCZK       = rows.reduce((s, r) => s + Number(r['Sale in CZK'] ?? 0), 0)
  const exemptSaleCZK      = rows.reduce((s, r) => r['Tax exemption'] === 'Yes' ? s + Number(r['Sale in CZK'] ?? 0) : s, 0)
  const taxableSaleCZK     = rows.reduce((s, r) => r['Tax exemption'] === 'No'  ? s + Number(r['Sale in CZK'] ?? 0) : s, 0)
  const taxBase    = Math.max(0, taxableSaleCZK - taxablePurchaseCZK)
  const mustDeclare = totalSaleCZK > 99999 && taxBase > 0

  // Pivot: summarise per Symbol
  const symbolPivot = useMemo(() => {
    const map = new Map<string, {
      symbol: string
      trades: number
      purchaseCZK: number
      saleCZK: number
      exemptPurchaseCZK: number
      exemptSaleCZK: number
      taxablePurchaseCZK: number
      taxableSaleCZK: number
    }>()
    for (const r of rows) {
      const sym = String(r['Symbol'] ?? '—')
      if (!map.has(sym)) map.set(sym, { symbol: sym, trades: 0, purchaseCZK: 0, saleCZK: 0, exemptPurchaseCZK: 0, exemptSaleCZK: 0, taxablePurchaseCZK: 0, taxableSaleCZK: 0 })
      const e = map.get(sym)!
      const p = Number(r['Purchase in CZK'] ?? 0)
      const s = Number(r['Sale in CZK'] ?? 0)
      const exempt = r['Tax exemption'] === 'Yes'
      e.trades++
      e.purchaseCZK += p
      e.saleCZK += s
      if (exempt) { e.exemptPurchaseCZK += p; e.exemptSaleCZK += s }
      else        { e.taxablePurchaseCZK += p; e.taxableSaleCZK += s }
    }
    return Array.from(map.values()).sort((a, b) => b.saleCZK - a.saleCZK)
  }, [rows])

  function formatCZK(value: number) {
    return new Intl.NumberFormat('cs-CZ', { style: 'currency', currency: 'CZK', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(value)
  }

  return (
    <div style={{ background: '#10151f', minHeight: '100vh', color: '#c9d1e0', fontFamily: "'Syne', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700&family=Fira+Code:wght@300;400;500&display=swap');

        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        :root {
          --bg:        #10151f;
          --bg-card:   #161d2b;
          --bg-row:    #1a2133;
          --border:    rgba(255,255,255,0.11);
          --border-hi: rgba(180,210,255,0.22);
          --accent:    #6aa3ff;
          --accent-dim:#2a4f8a;
          --green:     #3de0b0;
          --red:       #ff8090;
          --text-1:    #f0f4fc;
          --text-2:    #94a6be;
          --text-3:    #546478;
          --mono:      'Fira Code', monospace;
        }

        @keyframes fadeUp {
          from { opacity: 0; transform: translateY(16px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeIn {
          from { opacity: 0; }
          to   { opacity: 1; }
        }
        @keyframes pulse-ring {
          0%   { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(78,143,255,0.3); }
          70%  { transform: scale(1);    box-shadow: 0 0 0 10px rgba(78,143,255,0); }
          100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(78,143,255,0); }
        }
        @keyframes spin {
          to { transform: rotate(360deg); }
        }

        .appear-1 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) both; }
        .appear-2 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) 0.08s both; }
        .appear-3 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) 0.16s both; }
        .appear-4 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) 0.24s both; }

        /* ── Navbar ── */
        .navbar {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 0 40px;
          height: 56px;
          border-bottom: 1px solid var(--border);
          background: rgba(16,21,31,0.92);
          backdrop-filter: blur(12px);
          position: sticky;
          top: 0;
          z-index: 100;
        }
        .nav-brand {
          display: flex;
          align-items: center;
          gap: 10px;
          font-size: 13px;
          font-weight: 700;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: var(--text-1);
        }
        .nav-brand-dot {
          width: 7px; height: 7px;
          border-radius: 50%;
          background: var(--accent);
          box-shadow: 0 0 8px var(--accent);
        }
        .nav-file {
          display: flex;
          align-items: center;
          gap: 8px;
          font-family: var(--mono);
          font-size: 11px;
          color: var(--text-2);
          background: rgba(78,143,255,0.06);
          border: 1px solid rgba(78,143,255,0.12);
          border-radius: 5px;
          padding: 5px 12px;
          max-width: 400px;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        /* ── Upload button ── */
        .upload-btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 9px 20px;
          font-family: 'Syne', sans-serif;
          font-size: 12px;
          font-weight: 600;
          letter-spacing: 0.08em;
          text-transform: uppercase;
          color: var(--text-1);
          background: var(--bg-card);
          border: 1px solid var(--border-hi);
          border-radius: 6px;
          cursor: pointer;
          transition: all 0.2s;
          box-shadow: 0 2px 12px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.04);
        }
        .upload-btn:hover:not(:disabled) {
          background: #131b2e;
          border-color: rgba(78,143,255,0.4);
          box-shadow: 0 0 0 3px rgba(78,143,255,0.08), 0 2px 12px rgba(0,0,0,0.4);
          color: #fff;
        }
        .upload-btn:disabled { opacity: 0.4; cursor: not-allowed; }

        /* ── Empty state ── */
        .empty-wrap {
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          min-height: calc(100vh - 56px);
          gap: 20px;
          animation: fadeIn 0.6s ease both;
        }
        .empty-icon-ring {
          width: 72px; height: 72px;
          border-radius: 50%;
          border: 1px solid rgba(78,143,255,0.2);
          display: flex;
          align-items: center;
          justify-content: center;
          animation: pulse-ring 2.5s ease-in-out infinite;
          margin-bottom: 4px;
        }
        .empty-title {
          font-size: 15px;
          font-weight: 600;
          color: var(--text-1);
          letter-spacing: 0.04em;
        }
        .empty-sub {
          font-size: 12px;
          color: var(--text-3);
          letter-spacing: 0.04em;
          margin-top: -8px;
        }

        /* ── Content ── */
        .content { padding: 32px 40px; }

        /* ── Stat cards ── */
        .cards-row {
          display: flex;
          gap: 14px;
          flex-wrap: wrap;
          margin-bottom: 28px;
        }
        .card {
          background: var(--bg-card);
          border: 1px solid var(--border);
          border-radius: 10px;
          padding: 22px 24px 20px;
          min-width: 230px;
          flex: 1;
          max-width: 320px;
          position: relative;
          overflow: hidden;
          transition: border-color 0.2s, box-shadow 0.2s;
        }
        .card::before {
          content: '';
          position: absolute;
          top: 0; left: 0; right: 0;
          height: 1px;
          background: linear-gradient(90deg, transparent, rgba(78,143,255,0.3), transparent);
          opacity: 0;
          transition: opacity 0.3s;
        }
        .card:hover { border-color: var(--border-hi); box-shadow: 0 8px 32px rgba(0,0,0,0.4); }
        .card:hover::before { opacity: 1; }

        .card-eyebrow {
          font-size: 10px;
          font-weight: 600;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: var(--text-3);
          margin-bottom: 10px;
        }
        .card-value {
          font-family: var(--mono);
          font-size: 24px;
          font-weight: 400;
          color: var(--text-1);
          letter-spacing: -0.02em;
          margin-bottom: 18px;
          line-height: 1;
        }
        .card-sep {
          height: 1px;
          background: var(--border);
          margin-bottom: 14px;
        }
        .card-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 9px;
          font-size: 11.5px;
        }
        .card-row:last-child { margin-bottom: 0; }
        .card-row-label { color: var(--text-2); letter-spacing: 0.02em; }
        .num-green { font-family: var(--mono); color: var(--green); font-size: 12px; }
        .num-red   { font-family: var(--mono); color: var(--red);   font-size: 12px; }
        .num-blue  { font-family: var(--mono); color: var(--accent); font-size: 12px; }

        .declare-pill {
          display: inline-flex;
          align-items: center;
          padding: 3px 9px;
          border-radius: 4px;
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.1em;
          text-transform: uppercase;
        }
        .pill-yes { background: rgba(255,107,122,0.1); color: var(--red);   border: 1px solid rgba(255,107,122,0.25); }
        .pill-no  { background: rgba(45,212,160,0.1);  color: var(--green); border: 1px solid rgba(45,212,160,0.25); }

        /* ── Table ── */
        .table-shell {
          background: var(--bg-card);
          border: 1px solid var(--border);
          border-radius: 10px;
          overflow: hidden;
        }
        .table-topbar {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 14px 20px;
          border-bottom: 1px solid var(--border);
        }
        .table-topbar-left {
          display: flex;
          align-items: center;
          gap: 12px;
        }
        .table-label {
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: var(--text-3);
        }
        .table-count {
          font-family: var(--mono);
          font-size: 11px;
          color: var(--accent);
          background: rgba(78,143,255,0.08);
          border: 1px solid rgba(78,143,255,0.15);
          border-radius: 4px;
          padding: 2px 8px;
        }
        .export-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 7px 14px;
          font-family: 'Syne', sans-serif;
          font-size: 11px;
          font-weight: 600;
          letter-spacing: 0.06em;
          text-transform: uppercase;
          color: var(--text-2);
          background: transparent;
          border: 1px solid var(--border);
          border-radius: 6px;
          cursor: pointer;
          transition: all 0.2s;
        }
        .export-btn:hover {
          color: var(--text-1);
          border-color: var(--border-hi);
          background: rgba(255,255,255,0.03);
        }

        table { width: 100%; border-collapse: collapse; }
        thead tr { background: rgba(0,0,0,0.2); }
        thead th {
          font-family: 'Syne', sans-serif;
          font-size: 10px;
          font-weight: 600;
          letter-spacing: 0.1em;
          text-transform: uppercase;
          color: var(--text-3);
          padding: 10px 14px;
          text-align: right;
          white-space: nowrap;
          cursor: pointer;
          user-select: none;
          border-bottom: 1px solid var(--border);
          transition: color 0.15s;
        }
        thead th:first-child { text-align: left; }
        thead th:hover { color: var(--text-2); }
        thead th.th-active { color: var(--accent); }

        tbody tr {
          border-bottom: 1px solid rgba(255,255,255,0.03);
          transition: background 0.12s;
        }
        tbody tr:last-child { border-bottom: none; }
        tbody tr:nth-child(odd) { background: rgba(255,255,255,0.03); }
        tbody tr:hover { background: rgba(78,143,255,0.05) !important; }

        tbody td {
          padding: 9px 14px;
          font-family: var(--mono);
          font-size: 11.5px;
          color: var(--text-2);
          text-align: right;
          white-space: nowrap;
        }
        tbody td:first-child {
          text-align: left;
          font-family: 'Syne', sans-serif;
          font-size: 12px;
          font-weight: 500;
          color: var(--text-1);
        }
        tbody td:nth-child(2) {
          font-family: 'Syne', sans-serif;
          font-size: 12px;
          color: var(--text-1);
        }

        .spinner {
          width: 14px; height: 14px;
          border: 2px solid var(--border);
          border-top-color: var(--accent);
          border-radius: 50%;
          animation: spin 0.7s linear infinite;
        }

        /* ── Dashboard split layout ── */
        .dashboard-split {
          display: grid;
          grid-template-columns: auto 1fr;
          gap: 14px;
          margin-bottom: 28px;
          align-items: start;
        }
        .cards-col {
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .card { max-width: 260px; min-width: 220px; }

        /* ── Pivot panel ── */
        .pivot-shell {
          background: var(--bg-card);
          border: 1px solid var(--border);
          border-radius: 10px;
          overflow: hidden;
          height: 100%;
        }
        .pivot-topbar {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 14px 20px;
          border-bottom: 1px solid var(--border);
        }
        .pivot-topbar-left {
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .pivot-label {
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: var(--text-3);
        }
        .pivot-count {
          font-family: var(--mono);
          font-size: 11px;
          color: var(--accent);
          background: rgba(106,163,255,0.08);
          border: 1px solid rgba(106,163,255,0.15);
          border-radius: 4px;
          padding: 2px 8px;
        }
        .pivot-scroll { overflow-y: auto; max-height: 480px; }
        .pivot-table { width: 100%; border-collapse: collapse; }
        .pivot-table thead th {
          position: sticky;
          top: 0;
          background: rgba(0,0,0,0.35);
          font-family: 'Syne', sans-serif;
          font-size: 10px;
          font-weight: 600;
          letter-spacing: 0.1em;
          text-transform: uppercase;
          color: var(--text-3);
          padding: 9px 16px;
          text-align: right;
          white-space: nowrap;
          border-bottom: 1px solid var(--border);
          user-select: none;
          z-index: 1;
        }
        .pivot-table thead th:first-child { text-align: left; }
        .pivot-table tbody tr {
          border-bottom: 1px solid rgba(255,255,255,0.03);
          transition: background 0.12s;
        }
        .pivot-table tbody tr:last-child { border-bottom: none; }
        .pivot-table tbody tr:nth-child(odd) { background: rgba(255,255,255,0.02); }
        .pivot-table tbody tr:hover { background: rgba(106,163,255,0.05) !important; }
        .pivot-table tbody td {
          padding: 9px 16px;
          font-family: var(--mono);
          font-size: 11.5px;
          color: var(--text-2);
          text-align: right;
          white-space: nowrap;
        }
        .pivot-table tbody td:first-child {
          text-align: left;
          font-family: 'Syne', sans-serif;
          font-size: 12px;
          font-weight: 600;
          color: var(--text-1);
          letter-spacing: 0.02em;
        }
        .pivot-table tbody td:nth-child(2) {
          font-family: 'Syne', sans-serif;
          font-size: 11px;
          color: var(--text-3);
          text-align: center;
        }
        .pnl-pos { color: var(--green); }
        .pnl-neg { color: var(--red); }
        .pnl-zero { color: var(--text-3); }
      `}</style>

      {/* ── Navbar ── */}
      <nav className="navbar">
        <div className="nav-brand">
          <div className="nav-brand-dot" />
          XTB Tax
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          {fileName && (
            <div className="nav-file">
              <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
              </svg>
              {fileName}
            </div>
          )}
          <button className="upload-btn" onClick={triggerFileUpload} disabled={loading}>
            {loading
              ? <><div className="spinner" /> Processing</>
              : <><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg> Import</>
            }
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx" onChange={handleFileUpload} hidden />
        </div>
      </nav>

      {/* ── Empty state ── */}
      {!fileName && !loading && (
        <div className="empty-wrap">
          <div className="empty-icon-ring">
            <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="rgba(78,143,255,0.6)" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
              <line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
            </svg>
          </div>
          <div className="empty-title">Import your XTB statement</div>
          <div className="empty-sub">Upload a .xlsx export file to calculate your tax obligations</div>
          <button className="upload-btn" onClick={triggerFileUpload}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
            Choose File
          </button>
          {error && <p style={{ color: 'var(--red)', fontSize: 12, marginTop: 4 }}>{error}</p>}
        </div>
      )}

      {/* ── Main content ── */}
      {rows.length > 0 && (
        <div className="content">

          {error && <p style={{ color: 'var(--red)', fontSize: 12, marginBottom: 16 }}>{error}</p>}

          {/* Dashboard split: stat cards left, pivot right */}
          <div className="dashboard-split">

            {/* Left col — stat cards */}
            <div className="cards-col">

              <div className="card appear-1">
                <div className="card-eyebrow">Total Purchases</div>
                <div className="card-value">{formatCZK(totalPurchaseCZK)}</div>
                <div className="card-sep" />
                <div className="card-row">
                  <span className="card-row-label">Tax exempt</span>
                  <span className="num-green">{formatCZK(exemptPurchaseCZK)}</span>
                </div>
                <div className="card-row">
                  <span className="card-row-label">Taxable</span>
                  <span className="num-red">{formatCZK(taxablePurchaseCZK)}</span>
                </div>
              </div>

              <div className="card appear-2">
                <div className="card-eyebrow">Total Sales</div>
                <div className="card-value">{formatCZK(totalSaleCZK)}</div>
                <div className="card-sep" />
                <div className="card-row">
                  <span className="card-row-label">Tax exempt</span>
                  <span className="num-green">{formatCZK(exemptSaleCZK)}</span>
                </div>
                <div className="card-row">
                  <span className="card-row-label">Taxable</span>
                  <span className="num-red">{formatCZK(taxableSaleCZK)}</span>
                </div>
              </div>

              <div className="card appear-3">
                <div className="card-eyebrow">Tax Base</div>
                <div className="card-value">{formatCZK(taxBase)}</div>
                <div className="card-sep" />
                <div className="card-row">
                  <span className="card-row-label">Declare income?</span>
                  <span className={`declare-pill ${mustDeclare ? 'pill-yes' : 'pill-no'}`}>
                    {mustDeclare ? 'Yes' : 'No'}
                  </span>
                </div>
              </div>

            </div>

            {/* Right col — symbol pivot */}
            <div className="pivot-shell appear-2">
              <div className="pivot-topbar">
                <div className="pivot-topbar-left">
                  <span className="pivot-label">By Symbol</span>
                  <span className="pivot-count">{symbolPivot.length}</span>
                </div>
              </div>
              <div className="pivot-scroll">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th>Symbol</th>
                      <th style={{ textAlign: 'center' }}>Trades</th>
                      <th>Purchases CZK</th>
                      <th>Sales CZK</th>
                      <th>P&amp;L CZK</th>
                      <th>Taxable Buy</th>
                      <th>Taxable Sell</th>
                    </tr>
                  </thead>
                  <tbody>
                    {symbolPivot.map(s => {
                      const pnl = s.saleCZK - s.purchaseCZK
                      const pnlClass = pnl > 0 ? 'pnl-pos' : pnl < 0 ? 'pnl-neg' : 'pnl-zero'
                      return (
                        <tr key={s.symbol}>
                          <td>{s.symbol}</td>
                          <td>{s.trades}</td>
                          <td>{formatCZK(s.purchaseCZK)}</td>
                          <td>{formatCZK(s.saleCZK)}</td>
                          <td className={pnlClass}>{pnl >= 0 ? '+' : ''}{formatCZK(pnl)}</td>
                          <td className="num-red">{formatCZK(s.taxablePurchaseCZK)}</td>
                          <td className="num-red">{formatCZK(s.taxableSaleCZK)}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>

          </div>

          {/* Trades table */}
          <div className="table-shell appear-4">
            <div className="table-topbar">
              <div className="table-topbar-left">
                <span className="table-label">Closed Positions</span>
                <span className="table-count">{rows.length}</span>
              </div>
              <button className="export-btn" onClick={exportToExcel}>
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Export Excel
              </button>
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table>
                <thead>
                  <tr>
                    {visibleKeys.map(k => (
                      <th
                        key={k}
                        className={sortKey === k ? 'th-active' : ''}
                        onClick={() => {
                          if (sortKey === k) setSortDir(p => p === 'asc' ? 'desc' : 'asc')
                          else { setSortKey(k); setSortDir('asc') }
                        }}
                      >
                        {k}{sortKey && sortKey === k ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sortedRows.map((r, i) => (
                    <tr key={i}>
                      {visibleKeys.map(k => (
                        <td key={k}>
                          {k.toLowerCase().includes('time')
                            ? formatExcelDate(r[k]).split(' ')[0]
                            : String(r[k] ?? '')}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

        </div>
      )}
    </div>
  )
}
