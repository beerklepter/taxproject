'use client'

import { useState, useRef, useMemo, useEffect } from 'react'
import * as XLSX from 'xlsx'

const HIDDEN_COLUMNS = new Set(['Open origin', 'Close origin', 'SL', 'TP', 'Margin', 'Comment', '_openDisplay', '_closeDisplay', 'Position', 'Open price', 'Close price', 'Commission', 'Swap', 'Rollover'])

type TradeRow = Record<string, any>

function formatCZK(value: number) {
  return new Intl.NumberFormat('cs-CZ', { style: 'currency', currency: 'CZK', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(value)
}

const CACHE_VERSION = 3

export default function Home() {
  const [fileName, setFileName]         = useState('')
  const [rows, setRows]                 = useState<TradeRow[]>([])
  const [openRows, setOpenRows]           = useState<TradeRow[]>([])
  const [error, setError]               = useState('')
  const [page, setPage]                 = useState(0)
  const [loading, setLoading]           = useState(false)
  const [loadingMsg, setLoadingMsg]     = useState('')
  const [fxWarnings, setFxWarnings]     = useState<string[]>([])
  const [selectedYear, setSelectedYear] = useState<number>(0)
  const [sortKey, setSortKey]           = useState<string>('')
  const [sortDir, setSortDir]           = useState<'asc' | 'desc'>('asc')
  const [openSortKey, setOpenSortKey]   = useState<string>('')
  const [openSortDir, setOpenSortDir]   = useState<'asc' | 'desc'>('asc')
  const [lastImported, setLastImported] = useState<string>('')
  const [todayFx, setTodayFx]             = useState<number>(0)
  const [restored, setRestored]         = useState(false)
  const [activeView, setActiveView]     = useState<string>('overview')
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Fetch today's EUR/CZK rate when open positions are loaded
  useEffect(() => {
    if (!openRows.length) return
    const today = new Date()
    const todayStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`
    fetchFx(todayStr).then(rate => { if (rate) setTodayFx(rate as number) })
  }, [openRows])

  // Persist selectedYear whenever it changes
  useEffect(() => {
    if (!selectedYear) return
    localStorage.setItem('xtb_selected_year', String(selectedYear))
  }, [selectedYear])

  // Restore persisted data on mount
  useEffect(() => {
    try {
      const saved = localStorage.getItem('xtb_cache')
      if (saved) {
        const parsed = JSON.parse(saved)
        if (parsed.version !== CACHE_VERSION) { localStorage.removeItem('xtb_cache'); return }
        const { fileName: fn, rows: r, openRows: or, fxWarnings: w, lastImported: li } = parsed
        if (fn && Array.isArray(r) && r.length > 0) {
          setFileName(fn)
          setRows(r)
          setOpenRows(or ?? [])
          setFxWarnings(w ?? [])
          setLastImported(li ?? '')
          setRestored(true)
          // Restore selected year from its own key
          const savedYear = localStorage.getItem('xtb_selected_year')
          if (savedYear && Number(savedYear) > 0) {
            setSelectedYear(Number(savedYear))
          } else {
            const restoredYears = [...new Set((r as TradeRow[]).map((row: TradeRow) => row['_closeYear']).filter(Boolean) as number[])].sort((a: number, b: number) => b - a)
            setSelectedYear(restoredYears[0] ?? 0)
          }
        }
      }
    } catch { /* ignore corrupt cache */ }
  }, [])

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

  function getCloseYear(closeValue: any): number | null {
    if (typeof closeValue !== 'number') return null
    const date = XLSX.SSF.parse_date_code(closeValue)
    return date ? date.y : null
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
    fxCache.current.clear()
    setLoading(true)
    setLoadingMsg('Reading file…')
    setFileName(file.name)
    setRows([])
    setOpenRows([])
    setError('')
    setFxWarnings([])
    setSelectedYear(0)
    setLastImported('')
    setRestored(false)
    setPage(0)

    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheetName = workbook.SheetNames.find(name => name.toLowerCase().includes('closed'))
        if (!sheetName) { setError('No CLOSED POSITION sheet found'); setLoading(false); setLoadingMsg(''); return }
        const sheet = workbook.Sheets[sheetName]
        const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][]
        const KNOWN_HEADERS = new Set(['Symbol', 'Open time', 'Close time', 'Volume'])
        const headerIndex = raw.findIndex(row => Array.isArray(row) && row.filter(cell => KNOWN_HEADERS.has(String(cell).trim())).length >= 3)
        if (headerIndex === -1) { setError('Could not find header row — unexpected file format'); setLoading(false); setLoadingMsg(''); return }
        const headers = raw[headerIndex] as string[]
        const REQUIRED_COLUMNS = ['Symbol', 'Open time', 'Close time', 'Purchase value', 'Sale value']
        const missingColumns = REQUIRED_COLUMNS.filter(col => !headers.includes(col))
        if (missingColumns.length > 0) { setError(`Missing expected columns: ${missingColumns.join(', ')}`); setLoading(false); setLoadingMsg(''); return }
        const dataRows = raw.slice(headerIndex + 1)
        const cleaned = dataRows.filter(r => { const first = String(r?.[0] ?? '').toLowerCase().trim(); return first && first !== 'total' })

        const partials = cleaned.map(row => {
          const obj: TradeRow = {}
          headers.forEach((h, i) => { obj[h] = row[i] })
          obj['Tax exemption']  = getTaxExemption(obj['Open time'], obj['Close time'])
          obj['_closeYear']     = getCloseYear(obj['Close time'])
          obj['_openDate']      = excelDateToApiDate(obj['Open time'])
          obj['_closeDate']     = excelDateToApiDate(obj['Close time'])
          obj['_openDisplay']   = formatExcelDate(obj['Open time'])
          obj['_closeDisplay']  = formatExcelDate(obj['Close time'])
          return obj
        })

        // Point 3: progress message while fetching FX rates
        const uniqueDates = [...new Set(partials.flatMap(o => [o['_openDate'], o['_closeDate']]).filter(Boolean))]
        setLoadingMsg(`Fetching exchange rates… 0 / ${uniqueDates.length}`)

        let fetched = 0
        await Promise.all(uniqueDates.map(async date => {
          await fetchFx(date)
          fetched++
          setLoadingMsg(`Fetching exchange rates… ${fetched} / ${uniqueDates.length}`)
        }))

        setLoadingMsg('Processing trades…')

        // Point 5: collect dates where FX was missing
        const missingFxDates = new Set<string>()

        const json = await Promise.all(partials.map(async obj => {
          const openFx  = await fetchFx(obj['_openDate'])
          const closeFx = await fetchFx(obj['_closeDate'])

          if (!openFx  && obj['_openDate'])  missingFxDates.add(obj['_openDate'])
          if (!closeFx && obj['_closeDate']) missingFxDates.add(obj['_closeDate'])

          obj['Open date fx']    = openFx  ? (openFx  as number).toFixed(3) : ''
          obj['Close date fx']   = closeFx ? (closeFx as number).toFixed(3) : ''
          const purchase         = Number(obj['Purchase value'] ?? 0)
          const sale             = Number(obj['Sale value']     ?? 0)
          obj['Purchase in CZK'] = openFx  ? (purchase * (openFx  as number)).toFixed(2) : ''
          obj['Sale in CZK']     = closeFx ? (sale     * (closeFx as number)).toFixed(2) : ''
          delete obj['_openDate']; delete obj['_closeDate']
          return obj
        }))

        if (missingFxDates.size > 0) {
          setFxWarnings([...missingFxDates].sort())
        }

        // ── Parse open positions sheet ──
        let outerOpenJson: TradeRow[] = []
        const openSheetName = workbook.SheetNames.find(n => n.toUpperCase().startsWith('OPEN POSITION'))
        if (openSheetName) {
          const openSheet = workbook.Sheets[openSheetName]
          const openRaw = XLSX.utils.sheet_to_json(openSheet, { header: 1, defval: '' }) as any[][]

          // Find header row
          const OPEN_KNOWN = new Set(['Symbol', 'Open time', 'Volume', 'Open price'])
          const openHeaderIdx = openRaw.findIndex(row =>
            Array.isArray(row) && row.filter(c => OPEN_KNOWN.has(String(c).trim())).length >= 3
          )

          if (openHeaderIdx !== -1) {
            const openHeaders = openRaw[openHeaderIdx] as string[]
            const openDataRows = openRaw.slice(openHeaderIdx + 1).filter(r => {
              const hasData = r.some((c: any) => c !== '' && c !== null && c !== undefined)
              const firstNonEmpty = r.find((c: any) => c !== '' && c !== null && c !== undefined)
              const firstStr = String(firstNonEmpty ?? '').toLowerCase().trim()
              return hasData && firstStr !== 'total'
            })

            // Detect column offset — XTB open positions have a leading empty col in data rows
            const firstData = openDataRows[0] ?? []
            const offset = (firstData[0] === '' || firstData[0] === null || firstData[0] === undefined) ? 1 : 0

            const now = new Date()
            const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()) // date only, no time
            const openJson = openDataRows.map(row => {
              const obj: TradeRow = {}
              openHeaders.forEach((h, i) => { obj[h] = row[i + offset] })

              // Format open time display
              obj['_openDisplay'] = formatExcelDate(obj['Open time'])

              // Days held (open date to today)
              const openApiDate = excelDateToApiDate(obj['Open time'])
              obj['_openApiDate'] = openApiDate
              if (openApiDate) {
                const openDateObj = new Date(openApiDate) // already midnight since from YYYY-MM-DD string
                const diffMs = today.getTime() - openDateObj.getTime()
                obj['Days held'] = Math.round(diffMs / (1000 * 60 * 60 * 24))
              } else {
                obj['Days held'] = ''
              }

              // Tax exemption — based on open date vs today (3-year rule)
              if (openApiDate) {
                const openDateObj = new Date(openApiDate)
                const threeYearsLater = new Date(openDateObj.getFullYear() + 3, openDateObj.getMonth(), openDateObj.getDate() + 1)
                if (today >= threeYearsLater) {
                  obj['Tax exempt'] = 'Yes'
                  obj['Days to exempt'] = 0
                } else {
                  obj['Tax exempt'] = 'No'
                  obj['Days to exempt'] = Math.round((threeYearsLater.getTime() - today.getTime()) / (1000 * 60 * 60 * 24))
                }
              }

              return obj
            })

            // Clean up internal date key
            for (const obj of openJson) {
              delete obj['_openApiDate']
            }

            setOpenRows(openJson)
            outerOpenJson = openJson
          }
        }

        setSortKey('')
        setSortDir('asc')
        setOpenSortKey('')
        setOpenSortDir('asc')
        // Default to latest year (availableYears is sorted desc so first unique close year)
        const years = [...new Set(json.map((r: TradeRow) => r['_closeYear']).filter(Boolean) as number[])].sort((a, b) => b - a)
        setSelectedYear(years[0] ?? 0)

        // Persist to localStorage — catch quota errors gracefully
        const importedAt = new Date().toLocaleString('cs-CZ')
        try {
          const latestYear = years[0] ?? 0
          localStorage.setItem('xtb_cache', JSON.stringify({
            version: CACHE_VERSION,
            fileName: file.name,
            rows: json,
            openRows: outerOpenJson,
            fxWarnings: [...missingFxDates].sort(),
            lastImported: importedAt
          }))
          localStorage.setItem('xtb_selected_year', String(latestYear))
        } catch (e) {
          console.warn('localStorage quota exceeded — data not cached', e)
        }
        setLastImported(importedAt)
        setRestored(false)
        setRows(json)
      } catch (err) { console.error(err); setError('Failed to process file') }
      finally { setLoading(false); setLoadingMsg('') }
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

  // Point 1: available years derived from close dates
  const availableYears = useMemo(() => {
    const years = new Set<number>()
    for (const r of rows) { if (r['_closeYear']) years.add(r['_closeYear']) }
    return [...years].sort((a, b) => b - a)
  }, [rows])

  // Point 1: filter rows by selected year
  const filteredRows = useMemo(() => {
    if (!selectedYear) return rows
    return rows.filter(r => r['_closeYear'] === selectedYear)
  }, [rows, selectedYear])

  const visibleKeys = useMemo(() => filteredRows.length > 0 ? Object.keys(filteredRows[0]).filter(k => !HIDDEN_COLUMNS.has(k) && k !== '_closeYear') : [], [filteredRows])

  const sortedRows = useMemo(() => {
    if (!sortKey) return filteredRows
    return [...filteredRows].sort((a, b) => {
      const av = a[sortKey], bv = b[sortKey]
      const aNum = Number(av), bNum = Number(bv)
      const bothNumbers = !isNaN(aNum) && !isNaN(bNum)
      const result = bothNumbers ? aNum - bNum : String(av ?? '').localeCompare(String(bv ?? ''))
      return sortDir === 'asc' ? result : -result
    })
  }, [filteredRows, sortKey, sortDir])

  const sortedOpenRows = useMemo(() => {
    if (!openSortKey) return openRows
    return [...openRows].sort((a, b) => {
      const av = a[openSortKey], bv = b[openSortKey]
      const aNum = Number(av), bNum = Number(bv)
      const bothNumbers = !isNaN(aNum) && !isNaN(bNum)
      const result = bothNumbers ? aNum - bNum : String(av ?? '').localeCompare(String(bv ?? ''))
      return openSortDir === 'asc' ? result : -result
    })
  }, [openRows, openSortKey, openSortDir])

  const {
    totalPurchaseCZK, exemptPurchaseCZK, taxablePurchaseCZK,
    totalSaleCZK, exemptSaleCZK, taxableSaleCZK,
    taxBase, mustDeclare, taxOwed
  } = useMemo(() => {
    let totalPurchaseCZK = 0, exemptPurchaseCZK = 0, taxablePurchaseCZK = 0
    let totalSaleCZK = 0, exemptSaleCZK = 0, taxableSaleCZK = 0
    for (const r of filteredRows) {
      const p = Number(r['Purchase in CZK'] ?? 0)
      const s = Number(r['Sale in CZK'] ?? 0)
      totalPurchaseCZK += p; totalSaleCZK += s
      if (r['Tax exemption'] === 'Yes') { exemptPurchaseCZK += p; exemptSaleCZK += s }
      else                              { taxablePurchaseCZK += p; taxableSaleCZK += s }
    }
    const taxBase = Math.max(0, taxableSaleCZK - taxablePurchaseCZK)
    const mustDeclare = totalSaleCZK > 100000
    const taxOwed = mustDeclare && taxBase > 0
    return { taxOwed, totalPurchaseCZK, exemptPurchaseCZK, taxablePurchaseCZK, totalSaleCZK, exemptSaleCZK, taxableSaleCZK, taxBase, mustDeclare }
  }, [filteredRows])

  const symbolPivot = useMemo(() => {
    const map = new Map<string, { symbol: string; trades: number; exemptTrades: number; taxablePurchaseCZK: number; taxableSaleCZK: number }>()
    for (const r of filteredRows) {
      const sym = String(r['Symbol'] ?? '—')
      if (!map.has(sym)) map.set(sym, { symbol: sym, trades: 0, exemptTrades: 0, taxablePurchaseCZK: 0, taxableSaleCZK: 0 })
      const e = map.get(sym)!
      const p = Number(r['Purchase in CZK'] ?? 0)
      const s = Number(r['Sale in CZK'] ?? 0)
      const exempt = r['Tax exemption'] === 'Yes'
      e.trades++
      if (exempt) { e.exemptTrades++ }
      else        { e.taxablePurchaseCZK += p; e.taxableSaleCZK += s }
    }
    return Array.from(map.values()).sort((a, b) => b.taxableSaleCZK - a.taxableSaleCZK)
  }, [filteredRows])

  return (
    <div style={{ background: '#10151f', minHeight: '100vh', color: '#c9d1e0', fontFamily: "'Syne', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700&family=Fira+Code:wght@300;400;500&display=swap');

        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        :root {
          --bg:        #10151f;
          --bg-card:   #161d2b;
          --border:    rgba(255,255,255,0.11);
          --border-hi: rgba(180,210,255,0.22);
          --accent:    #6aa3ff;
          --green:     #3de0b0;
          --red:       #ff8090;
          --amber:     #f5a623;
          --text-1:    #f0f4fc;
          --text-2:    #94a6be;
          --text-3:    #546478;
          --mono:      'Fira Code', monospace;
        }

        @keyframes fadeUp  { from { opacity:0; transform:translateY(16px); } to { opacity:1; transform:translateY(0); } }
        @keyframes fadeIn  { from { opacity:0; } to { opacity:1; } }
        @keyframes pulse-ring {
          0%   { transform:scale(0.95); box-shadow:0 0 0 0 rgba(78,143,255,0.3); }
          70%  { transform:scale(1);    box-shadow:0 0 0 10px rgba(78,143,255,0); }
          100% { transform:scale(0.95); box-shadow:0 0 0 0 rgba(78,143,255,0); }
        }
        @keyframes spin { to { transform:rotate(360deg); } }

        .appear-1 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) both; }
        .appear-2 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) 0.08s both; }
        .appear-3 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) 0.16s both; }
        .appear-4 { animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) 0.24s both; }

        /* ── Navbar ── */
        .navbar {
          display: flex; align-items: center; justify-content: space-between;
          padding: 0 40px; height: 56px;
          border-bottom: 1px solid var(--border);
          background: rgba(16,21,31,0.92);
          backdrop-filter: blur(12px);
          position: sticky; top: 0; z-index: 100;
        }
        .nav-brand { display:flex; align-items:center; gap:10px; font-size:13px; font-weight:700; letter-spacing:0.12em; text-transform:uppercase; color:var(--text-1); }
        .nav-brand-dot { width:7px; height:7px; border-radius:50%; background:var(--accent); box-shadow:0 0 8px var(--accent); }
        .nav-right { display:flex; align-items:center; gap:12px; }
        .nav-file {
          display:flex; align-items:center; gap:8px;
          font-family:var(--mono); font-size:11px; color:var(--text-2);
          background:rgba(78,143,255,0.06); border:1px solid rgba(78,143,255,0.12);
          border-radius:5px; padding:5px 12px;
          max-width:400px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
        }

        /* ── Year filter ── */
        .year-filter {
          display: flex; align-items: center; gap: 6px;
        }
        .year-label {
          font-size: 10px; font-weight: 600; letter-spacing: 0.1em;
          text-transform: uppercase; color: var(--text-3);
        }
        .year-btn {
          padding: 5px 11px;
          font-family: 'Syne', sans-serif; font-size: 11px; font-weight: 600;
          letter-spacing: 0.04em;
          color: var(--text-2);
          background: transparent;
          border: 1px solid var(--border);
          border-radius: 5px;
          cursor: pointer;
          transition: all 0.15s;
        }
        .year-btn:hover { color: var(--text-1); border-color: var(--border-hi); }
        .year-btn.active {
          color: var(--accent);
          background: rgba(106,163,255,0.1);
          border-color: rgba(106,163,255,0.3);
        }

        /* ── Upload button ── */
        .upload-btn {
          display:inline-flex; align-items:center; gap:8px; padding:9px 20px;
          font-family:'Syne',sans-serif; font-size:12px; font-weight:600; letter-spacing:0.08em; text-transform:uppercase;
          color:var(--text-1); background:var(--bg-card); border:1px solid var(--border-hi); border-radius:6px;
          cursor:pointer; transition:all 0.2s;
          box-shadow:0 2px 12px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.04);
        }
        .upload-btn:hover:not(:disabled) { background:#131b2e; border-color:rgba(78,143,255,0.4); box-shadow:0 0 0 3px rgba(78,143,255,0.08),0 2px 12px rgba(0,0,0,0.4); color:#fff; }
        .upload-btn:disabled { opacity:0.4; cursor:not-allowed; }

        /* ── Empty state ── */
        .empty-wrap { display:flex; flex-direction:column; align-items:center; justify-content:center; min-height:calc(100vh - 56px); gap:20px; animation:fadeIn 0.6s ease both; }
        .empty-icon-ring {
          width: 88px; height: 88px;
          border-radius: 50%;
          border: 1px solid rgba(78,143,255,0.25);
          display: flex; align-items: center; justify-content: center;
          animation: pulse-ring 2.5s ease-in-out infinite;
          margin-bottom: 4px;
          cursor: pointer;
          transition: border-color 0.2s, background 0.2s, box-shadow 0.2s, transform 0.15s;
          background: rgba(78,143,255,0.04);
        }
        .empty-icon-ring:hover {
          border-color: rgba(78,143,255,0.6);
          background: rgba(78,143,255,0.1);
          box-shadow: 0 0 0 6px rgba(78,143,255,0.08);
          transform: scale(1.05);
          animation: none;
        }
        .empty-title { font-size:15px; font-weight:600; color:var(--text-1); letter-spacing:0.04em; }
        .empty-sub { font-size:12px; color:var(--text-3); letter-spacing:0.04em; margin-top:-8px; }

        /* ── Loading message ── */
        .loading-msg {
          display: inline-flex; align-items: center; gap: 8px;
          font-family: var(--mono); font-size: 11px; color: var(--text-2);
        }

        /* ── FX warning banner ── */
        .fx-warning {
          display: flex; align-items: flex-start; gap: 12px;
          background: rgba(245,166,35,0.07);
          border: 1px solid rgba(245,166,35,0.2);
          border-radius: 8px;
          padding: 12px 16px;
          margin-bottom: 20px;
          font-size: 12px;
          color: var(--amber);
          line-height: 1.6;
        }
        .fx-warning-icon { flex-shrink: 0; margin-top: 1px; }
        .fx-warning-dates {
          font-family: var(--mono); font-size: 11px;
          color: rgba(245,166,35,0.7);
          margin-top: 4px;
        }
        .fx-warning-close {
          margin-left: auto; flex-shrink: 0;
          background: none; border: none; cursor: pointer;
          color: rgba(245,166,35,0.5); font-size: 16px; line-height: 1;
          padding: 0 0 0 8px;
          transition: color 0.15s;
        }
        .fx-warning-close:hover { color: var(--amber); }

        /* ── Content ── */
        /* .content replaced by .main-panel */

        /* ── Stat cards ── */
        .card {
          background:var(--bg-card); border:1px solid var(--border); border-radius:10px;
          padding:22px 24px 20px; min-width:220px; max-width:260px;
          position:relative; overflow:hidden; transition:border-color 0.2s, box-shadow 0.2s;
        }
        .card::before { content:''; position:absolute; top:0; left:0; right:0; height:1px; background:linear-gradient(90deg,transparent,rgba(78,143,255,0.3),transparent); opacity:0; transition:opacity 0.3s; }
        .card:hover { border-color:var(--border-hi); box-shadow:0 8px 32px rgba(0,0,0,0.4); }
        .card:hover::before { opacity:1; }
        .card-eyebrow { font-size:10px; font-weight:600; letter-spacing:0.12em; text-transform:uppercase; color:var(--text-3); margin-bottom:10px; }
        .card-value { font-family:var(--mono); font-size:24px; font-weight:400; color:var(--text-1); letter-spacing:-0.02em; margin-bottom:18px; line-height:1; }
        .card-sep { height:1px; background:var(--border); margin-bottom:14px; }
        .card-row { display:flex; justify-content:space-between; align-items:center; margin-bottom:9px; font-size:11.5px; }
        .card-row:last-child { margin-bottom:0; }
        .card-row-label { color:var(--text-2); }
        .num-green { font-family:var(--mono); color:var(--green); font-size:12px; }
        .num-red   { font-family:var(--mono); color:var(--red);   font-size:12px; }
        .declare-pill { display:inline-flex; align-items:center; padding:3px 9px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:0.1em; text-transform:uppercase; }
        .pill-yes { background:rgba(255,128,144,0.1); color:var(--red);   border:1px solid rgba(255,128,144,0.25); }
        .pill-no  { background:rgba(61,224,176,0.1);  color:var(--green); border:1px solid rgba(61,224,176,0.25); }

        /* ── Dashboard split ── */
        .dashboard-split { display:grid; grid-template-columns:auto 1fr; gap:14px; margin-bottom:28px; align-items:start; }
        .cards-col { display:flex; flex-direction:column; gap:14px; }

        /* ── Table shell ── */
        .table-shell { background:var(--bg-card); border:1px solid var(--border); border-radius:10px; overflow:hidden; }
        .table-topbar { display:flex; align-items:center; justify-content:space-between; padding:14px 20px; border-bottom:1px solid var(--border); }
        .table-topbar-left { display:flex; align-items:center; gap:12px; }
        .table-label { font-size:10px; font-weight:700; letter-spacing:0.12em; text-transform:uppercase; color:var(--text-3); }
        .table-count { font-family:var(--mono); font-size:11px; color:var(--accent); background:rgba(78,143,255,0.08); border:1px solid rgba(78,143,255,0.15); border-radius:4px; padding:2px 8px; }
        .export-btn { display:inline-flex; align-items:center; gap:6px; padding:7px 14px; font-family:'Syne',sans-serif; font-size:11px; font-weight:600; letter-spacing:0.06em; text-transform:uppercase; color:var(--text-2); background:transparent; border:1px solid var(--border); border-radius:6px; cursor:pointer; transition:all 0.2s; }
        .export-btn:hover { color:var(--text-1); border-color:var(--border-hi); background:rgba(255,255,255,0.03); }

        table { width:100%; border-collapse:collapse; }
        thead tr { background:rgba(0,0,0,0.2); }
        thead th { font-family:'Syne',sans-serif; font-size:10px; font-weight:600; letter-spacing:0.1em; text-transform:uppercase; color:var(--text-3); padding:9px 10px; text-align:right; white-space:nowrap; cursor:pointer; user-select:none; border-bottom:1px solid var(--border); transition:color 0.15s; }
        thead th:first-child { text-align:left; }
        thead th:hover { color:var(--text-2); }
        thead th.th-active { color:var(--accent); }
        tbody tr { border-bottom:1px solid rgba(255,255,255,0.03); transition:background 0.12s; }
        tbody tr:last-child { border-bottom:none; }
        tbody tr:nth-child(odd) { background:rgba(255,255,255,0.03); }
        tbody tr:hover { background:rgba(78,143,255,0.05) !important; }
        tbody td { padding:8px 10px; font-family:var(--mono); font-size:11px; color:var(--text-2); text-align:right; white-space:nowrap; }
        tbody td:first-child { text-align:left; font-family:'Syne',sans-serif; font-size:12px; font-weight:500; color:var(--text-1); }
        tbody td:nth-child(2) { font-family:'Syne',sans-serif; font-size:12px; color:var(--text-1); }

        .spinner { width:14px; height:14px; border:2px solid var(--border); border-top-color:var(--accent); border-radius:50%; animation:spin 0.7s linear infinite; }

        /* ── Pivot panel ── */
        .pivot-shell { background:var(--bg-card); border:1px solid var(--border); border-radius:10px; overflow:hidden; height:100%; }
        .pivot-topbar { display:flex; align-items:center; justify-content:space-between; padding:14px 20px; border-bottom:1px solid var(--border); }
        .pivot-topbar-left { display:flex; align-items:center; gap:10px; }
        .pivot-label { font-size:10px; font-weight:700; letter-spacing:0.12em; text-transform:uppercase; color:var(--text-3); }
        .pivot-count { font-family:var(--mono); font-size:11px; color:var(--accent); background:rgba(106,163,255,0.08); border:1px solid rgba(106,163,255,0.15); border-radius:4px; padding:2px 8px; }
        .pivot-scroll { overflow-y:auto; }
        .pivot-table { width:100%; border-collapse:collapse; }
        .pivot-table thead th { position:sticky; top:0; background:rgba(0,0,0,0.35); font-family:'Syne',sans-serif; font-size:10px; font-weight:600; letter-spacing:0.1em; text-transform:uppercase; color:var(--text-3); padding:9px 16px; text-align:right; white-space:nowrap; border-bottom:1px solid var(--border); user-select:none; z-index:1; }
        .pivot-table thead th:first-child { text-align:left; }
        .pivot-table tbody tr { border-bottom:1px solid rgba(255,255,255,0.03); transition:background 0.12s; }
        .pivot-table tbody tr:last-child { border-bottom:none; }
        .pivot-table tbody tr:nth-child(odd) { background:rgba(255,255,255,0.02); }
        .pivot-table tbody tr:hover { background:rgba(106,163,255,0.05) !important; }
        .pivot-table tbody td { padding:9px 16px; font-family:var(--mono); font-size:11.5px; color:var(--text-2); text-align:right; white-space:nowrap; }
        .pivot-table tbody td:first-child { text-align:left; font-family:'Syne',sans-serif; font-size:12px; font-weight:600; color:var(--text-1); letter-spacing:0.02em; }
        .pivot-table tbody td:nth-child(2) { font-family:'Syne',sans-serif; font-size:11px; color:var(--text-3); text-align:center; }
        .pnl-pos { color:var(--green); }
        .pnl-neg { color:var(--red); }
        .pnl-zero { color:var(--text-3); }

        /* ── Scrollbars — let browser handle show/hide natively ── */
        .pivot-scroll, .table-scroll { scrollbar-width:thin; scrollbar-color:rgba(106,163,255,0.25) transparent; }

        /* ── App shell with sidebar ── */
        .app-shell {
          display: flex;
          min-height: calc(100vh - 56px);
        }

        /* ── Sidebar ── */
        .sidebar {
          width: 210px;
          flex-shrink: 0;
          background: rgba(12,16,26,0.7);
          border-right: 1px solid var(--border);
          padding: 20px 0;
          display: flex;
          flex-direction: column;
          gap: 4px;
          position: sticky;
          top: 56px;
          height: calc(100vh - 56px);
          overflow-y: auto;
          scrollbar-width: none;
        }
        .sidebar::-webkit-scrollbar { display: none; }

        .sidebar-section {
          padding: 0 12px;
          margin-bottom: 4px;
        }
        .sidebar-section-label {
          font-size: 9px;
          font-weight: 700;
          letter-spacing: 0.14em;
          text-transform: uppercase;
          color: var(--text-3);
          padding: 10px 8px 6px;
        }

        .sidebar-item {
          display: flex;
          align-items: center;
          gap: 9px;
          padding: 8px 10px;
          border-radius: 7px;
          font-size: 12.5px;
          font-weight: 500;
          color: var(--text-2);
          cursor: pointer;
          transition: background 0.15s, color 0.15s;
          border: none;
          background: none;
          width: 100%;
          text-align: left;
          font-family: 'Syne', sans-serif;
          letter-spacing: 0.01em;
          user-select: none;
        }
        .sidebar-item:hover:not(.sidebar-item-disabled) {
          background: rgba(255,255,255,0.04);
          color: var(--text-1);
        }
        .sidebar-item.active {
          background: rgba(106,163,255,0.1);
          color: var(--accent);
        }
        .sidebar-item.active svg { stroke: var(--accent); }
        .sidebar-item-disabled {
          opacity: 0.35;
          cursor: not-allowed;
        }
        .sidebar-item svg {
          flex-shrink: 0;
          stroke: var(--text-3);
          transition: stroke 0.15s;
        }
        .sidebar-item:hover:not(.sidebar-item-disabled) svg {
          stroke: var(--text-2);
        }
        .sidebar-soon {
          margin-left: auto;
          font-size: 9px;
          font-weight: 700;
          letter-spacing: 0.08em;
          text-transform: uppercase;
          color: var(--text-3);
          background: rgba(255,255,255,0.05);
          border-radius: 3px;
          padding: 2px 5px;
          font-family: 'Syne', sans-serif;
        }
        .sidebar-divider {
          height: 1px;
          background: var(--border);
          margin: 8px 12px;
        }

        /* ── Main panel next to sidebar ── */
        .main-panel {
          flex: 1;
          min-width: 0;
          padding: 32px 36px;
        }
      `}</style>

      <input ref={fileInputRef} type="file" accept=".xlsx" onChange={handleFileUpload} hidden />

      {/* ── Navbar ── */}
      <nav className="navbar">
        <div className="nav-brand">
          <div className="nav-brand-dot" />
          XTB Tax
        </div>
        <div className="nav-right">

          {fileName && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div className="nav-file">
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                </svg>
                {fileName}
              </div>
              {lastImported && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ fontFamily: 'var(--mono)', fontSize: 10, color: 'var(--text-3)' }}>
                    {lastImported}
                  </span>

                </div>
              )}
            </div>
          )}
          {/* Import button — only shown after a file is loaded */}
          {fileName && (
            loading && loadingMsg
              ? <div className="loading-msg"><div className="spinner" />{loadingMsg}</div>
              : <button
                  onClick={triggerFileUpload}
                  disabled={loading}
                  style={{
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                    padding: '7px 12px',
                    fontFamily: 'Syne, sans-serif', fontSize: 11, fontWeight: 600,
                    letterSpacing: '0.06em', textTransform: 'uppercase',
                    color: 'var(--text-2)',
                    background: 'transparent',
                    border: '1px solid var(--border)',
                    borderRadius: 6,
                    cursor: 'pointer',
                    transition: 'all 0.2s',
                  }}
                  onMouseEnter={e => { e.currentTarget.style.color = 'var(--text-1)'; e.currentTarget.style.borderColor = 'var(--border-hi)' }}
                  onMouseLeave={e => { e.currentTarget.style.color = 'var(--text-2)'; e.currentTarget.style.borderColor = 'var(--border)' }}
                >
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                  Import New
                </button>
          )}
          {/* Close button — rightmost, only shown when file is loaded */}
          {fileName && !loading && (
            <button
              title="Close file and return to start"
              onClick={() => { localStorage.removeItem('xtb_cache'); localStorage.removeItem('xtb_selected_year'); setFileName(''); setRows([]); setOpenRows([]); setLastImported(''); setRestored(false); setFxWarnings([]); setSelectedYear(0); setActiveView('overview') }}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: 6,
                padding: '7px 12px',
                fontFamily: 'Syne, sans-serif', fontSize: 11, fontWeight: 600,
                letterSpacing: '0.06em', textTransform: 'uppercase',
                color: 'var(--text-3)',
                background: 'transparent',
                border: '1px solid var(--border)',
                borderRadius: 6,
                cursor: 'pointer',
                transition: 'all 0.2s',
              }}
              onMouseEnter={e => { e.currentTarget.style.color = 'var(--red)'; e.currentTarget.style.borderColor = 'rgba(255,128,144,0.3)' }}
              onMouseLeave={e => { e.currentTarget.style.color = 'var(--text-3)'; e.currentTarget.style.borderColor = 'var(--border)' }}
            >
              <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
              </svg>
              Close
            </button>
          )}
        </div>
      </nav>

      <div className="app-shell">

        {/* ── Sidebar ── */}
        {rows.length > 0 && (
          <aside className="sidebar">
            <div className="sidebar-section">
              <div className="sidebar-section-label">Trades</div>

              <button className={`sidebar-item ${activeView === 'overview' ? 'active' : ''}`} onClick={() => setActiveView('overview')}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
                </svg>
                Closed Trades
              </button>

              <button className={`sidebar-item ${activeView === 'open' ? 'active' : ''}`} onClick={() => setActiveView('open')}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
                </svg>
                Open Positions
              </button>
            </div>
          </aside>
        )}

        {/* ── Empty state (no sidebar) ── */}
        {!fileName && !loading && (
          <div className="empty-wrap" style={{ flex: 1 }}>
            <div className="empty-icon-ring" onClick={triggerFileUpload} title="Click to import file">
              <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="rgba(78,143,255,0.7)" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                <line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
              </svg>
            </div>
            <div className="empty-title">Import your XTB statement</div>
            <div className="empty-sub">Upload a .xlsx export file to calculate your tax obligations</div>
            {error && <p style={{ color: 'var(--red)', fontSize: 12, marginTop: 4 }}>{error}</p>}
          </div>
        )}

        {/* ── Main content ── */}
        {rows.length > 0 && (
          <div className="main-panel">

            {error && <p style={{ color: 'var(--red)', fontSize: 12, marginBottom: 16 }}>{error}</p>}

            {/* FX warning banner — always visible */}
          {fxWarnings.length > 0 && (
            <div className="fx-warning">
              <svg className="fx-warning-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
              </svg>
              <div>
                <div>Exchange rates could not be fetched for <strong>{fxWarnings.length} date{fxWarnings.length > 1 ? 's' : ''}</strong>. Trades on these dates have empty CZK values and are excluded from all totals.</div>
                <div className="fx-warning-dates">{fxWarnings.join(' · ')}</div>
              </div>
              <button className="fx-warning-close" onClick={() => setFxWarnings([])}>×</button>
            </div>
          )}

          {/* ── Overview: cards + pivot ── */}
          {activeView === 'overview' && (
          <>
          {availableYears.length > 0 && (
            <div className="year-filter" style={{ marginBottom: 20 }}>
              <span className="year-label">Year</span>
              {availableYears.map(y => (
                <button key={y} className={`year-btn ${selectedYear === y ? 'active' : ''}`} onClick={() => setSelectedYear(y)}>{y}</button>
              ))}
            </div>
          )}
          <div className="dashboard-split">
            <div className="cards-col">
              <div className="card appear-1">
                <div className="card-eyebrow">Total Purchases</div>
                <div className="card-value">{formatCZK(totalPurchaseCZK)}</div>
                <div className="card-sep" />
                <div className="card-row"><span className="card-row-label">Tax exempt</span><span className="num-green">{formatCZK(exemptPurchaseCZK)}</span></div>
                <div className="card-row"><span className="card-row-label">Taxable</span><span className="num-red">{formatCZK(taxablePurchaseCZK)}</span></div>
              </div>
              <div className="card appear-2">
                <div className="card-eyebrow">Total Sales</div>
                <div className="card-value">{formatCZK(totalSaleCZK)}</div>
                <div className="card-sep" />
                <div className="card-row"><span className="card-row-label">Tax exempt</span><span className="num-green">{formatCZK(exemptSaleCZK)}</span></div>
                <div className="card-row"><span className="card-row-label">Taxable</span><span className="num-red">{formatCZK(taxableSaleCZK)}</span></div>
              </div>
              <div className="card appear-3">
                <div className="card-eyebrow">Tax Base</div>
                <div className="card-value">{formatCZK(taxBase)}</div>
                <div className="card-sep" />
                <div className="card-row">
                  <span className="card-row-label">Declare income?</span>
                  <span className={`declare-pill ${mustDeclare ? 'pill-yes' : 'pill-no'}`}>{mustDeclare ? 'Yes' : 'No'}</span>
                </div>
                <div className="card-row">
                  <span className="card-row-label">Tax owed?</span>
                  <span className={`declare-pill ${taxOwed ? 'pill-yes' : 'pill-no'}`}>{taxOwed ? 'Yes' : 'No'}</span>
                </div>
              </div>
              <div className="card appear-3">
                {(() => {
                  const remaining = Math.max(0, taxablePurchaseCZK - taxableSaleCZK)
                  const pct = taxablePurchaseCZK > 0 ? Math.min(100, (taxableSaleCZK / taxablePurchaseCZK) * 100) : 0
                  const barColor = remaining <= 0 ? 'var(--red)' : pct >= 80 ? 'var(--amber)' : 'var(--green)'
                  return (
                    <>
                      <div className="card-eyebrow">Sellable Budget</div>
                      <div className="card-value" style={{ color: remaining <= 0 ? 'var(--red)' : 'var(--text-1)' }}>{formatCZK(remaining)}</div>
                      <div style={{ fontSize: 10, color: 'var(--text-3)', marginTop: -12, marginBottom: 14 }}>
                        before tax becomes due
                      </div>
                      <div style={{ height: 4, borderRadius: 99, background: 'rgba(255,255,255,0.06)', marginBottom: 14, overflow: 'hidden' }}>
                        <div style={{ height: '100%', width: `${pct}%`, borderRadius: 99, background: barColor, transition: 'width 0.4s ease' }} />
                      </div>
                      <div className="card-sep" />
                      <div className="card-row">
                        <span className="card-row-label">Taxable sales</span>
                        <span style={{ fontFamily: 'var(--mono)', fontSize: 12, color: barColor }}>{formatCZK(taxableSaleCZK)}</span>
                      </div>
                      <div className="card-row">
                        <span className="card-row-label">Taxable purchases</span>
                        <span style={{ fontFamily: 'var(--mono)', fontSize: 12, color: 'var(--text-2)' }}>{formatCZK(taxablePurchaseCZK)}</span>
                      </div>
                    </>
                  )
                })()}
              </div>
            </div>

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
                      <th style={{ textAlign:'center' }}>Trades</th>
                      <th>Taxable Buy</th>
                      <th>Taxable Sell</th>
                      <th>Tax Base</th>
                      <th style={{ textAlign:'center' }}>Exempt</th>
                    </tr>
                  </thead>
                  <tbody>
                    {symbolPivot.map(s => {
                      const taxBase = Math.max(0, s.taxableSaleCZK - s.taxablePurchaseCZK)
                      const taxBaseClass = taxBase > 0 ? 'pnl-pos' : 'pnl-zero'
                      return (
                        <tr key={s.symbol}>
                          <td>{s.symbol}</td>
                          <td style={{ textAlign:'center', fontFamily:'var(--mono)', color:'var(--text-3)' }}>{s.trades}</td>
                          <td className="num-red">{formatCZK(s.taxablePurchaseCZK)}</td>
                          <td className="num-red">{formatCZK(s.taxableSaleCZK)}</td>
                          <td className={taxBaseClass}>{taxBase > 0 ? '+' : ''}{formatCZK(taxBase)}</td>
                          <td style={{ textAlign:'center', fontFamily:'var(--mono)', fontSize:11, color: s.exemptTrades > 0 ? 'var(--green)' : 'var(--text-3)' }}>{s.exemptTrades} / {s.trades}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          </>
          )}

          {/* ── Tax Summary view: cards only ── */}
          {activeView === 'tax' && (
            <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap' }}>
              <div className="card appear-1" style={{ maxWidth: 300 }}>
                <div className="card-eyebrow">Total Purchases</div>
                <div className="card-value">{formatCZK(totalPurchaseCZK)}</div>
                <div className="card-sep" />
                <div className="card-row"><span className="card-row-label">Tax exempt</span><span className="num-green">{formatCZK(exemptPurchaseCZK)}</span></div>
                <div className="card-row"><span className="card-row-label">Taxable</span><span className="num-red">{formatCZK(taxablePurchaseCZK)}</span></div>
              </div>
              <div className="card appear-2" style={{ maxWidth: 300 }}>
                <div className="card-eyebrow">Total Sales</div>
                <div className="card-value">{formatCZK(totalSaleCZK)}</div>
                <div className="card-sep" />
                <div className="card-row"><span className="card-row-label">Tax exempt</span><span className="num-green">{formatCZK(exemptSaleCZK)}</span></div>
                <div className="card-row"><span className="card-row-label">Taxable</span><span className="num-red">{formatCZK(taxableSaleCZK)}</span></div>
              </div>
              <div className="card appear-3" style={{ maxWidth: 300 }}>
                <div className="card-eyebrow">Tax Base</div>
                <div className="card-value">{formatCZK(taxBase)}</div>
                <div className="card-sep" />
                <div className="card-row">
                  <span className="card-row-label">Declare income?</span>
                  <span className={`declare-pill ${mustDeclare ? 'pill-yes' : 'pill-no'}`}>{mustDeclare ? 'Yes' : 'No'}</span>
                </div>
              </div>
            </div>
          )}

          {/* ── By Symbol view: full-width pivot ── */}
          {activeView === 'symbols' && (
            <div className="pivot-shell appear-1">
              <div className="pivot-topbar">
                <div className="pivot-topbar-left">
                  <span className="pivot-label">By Symbol</span>
                  <span className="pivot-count">{symbolPivot.length}</span>
                </div>
              </div>
              <div style={{ overflowY: 'auto' }}>
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th>Symbol</th>
                      <th style={{ textAlign:'center' }}>Trades</th>
                      <th>Taxable Buy</th>
                      <th>Taxable Sell</th>
                      <th>Tax Base</th>
                      <th style={{ textAlign:'center' }}>Exempt</th>
                    </tr>
                  </thead>
                  <tbody>
                    {symbolPivot.map(s => {
                      const taxBase = Math.max(0, s.taxableSaleCZK - s.taxablePurchaseCZK)
                      const taxBaseClass = taxBase > 0 ? 'pnl-pos' : 'pnl-zero'
                      return (
                        <tr key={s.symbol}>
                          <td>{s.symbol}</td>
                          <td style={{ textAlign:'center', fontFamily:'var(--mono)', color:'var(--text-3)' }}>{s.trades}</td>
                          <td className="num-red">{formatCZK(s.taxablePurchaseCZK)}</td>
                          <td className="num-red">{formatCZK(s.taxableSaleCZK)}</td>
                          <td className={taxBaseClass}>{taxBase > 0 ? '+' : ''}{formatCZK(taxBase)}</td>
                          <td style={{ textAlign:'center', fontFamily:'var(--mono)', fontSize:11, color: s.exemptTrades > 0 ? 'var(--green)' : 'var(--text-3)' }}>{s.exemptTrades} / {s.trades}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ── Open Positions view ── */}
          {activeView === 'open' && (
            <div>
              {openRows.length === 0 ? (
                <div style={{ color: 'var(--text-3)', fontSize: 13, padding: '40px 0', textAlign: 'center' }}>
                  No open positions found in this file.
                </div>
              ) : (
                <>
                  {/* Card + symbol table side by side */}
                  <div style={{ display: 'flex', gap: 14, alignItems: 'flex-start', marginBottom: 24 }}>
                    <div className="card appear-1" style={{ minWidth: 220, flexShrink: 0 }}>
                      <div style={{ fontSize: 13, fontWeight: 700, letterSpacing: '0.04em', color: 'var(--text-1)', marginBottom: 14, paddingBottom: 10, borderBottom: '1px solid var(--border)' }}>Summary</div>
                      <div className="card-sep" style={{ marginTop: 10 }} />
                      <div style={{ marginBottom: 6 }}>
                        <div style={{ fontSize: 10, color: 'var(--text-3)', fontWeight: 600, letterSpacing: '0.08em', textTransform: 'uppercase', marginBottom: 4 }}>Total purchase value</div>
                        <div className="card-value" style={{ fontSize: 18, marginBottom: 0, textAlign: 'right' }}>{openRows.reduce((s, r) => s + Number(r['Purchase value'] ?? 0), 0).toFixed(2)} EUR</div>
                      </div>
                      <div className="card-sep" style={{ marginTop: 14 }} />
                      <div style={{ marginBottom: 14 }}>
                        <div style={{ fontSize: 10, color: 'var(--text-3)', fontWeight: 600, letterSpacing: '0.08em', textTransform: 'uppercase', marginBottom: 4 }}>Unrealised P/L</div>
                        <div style={{ fontFamily: 'var(--mono)', fontSize: 15, fontWeight: 400, textAlign: 'right' }} className={openRows.reduce((s, r) => s + Number(r['Gross P/L'] ?? 0), 0) >= 0 ? 'num-green' : 'num-red'}>
                          {openRows.reduce((s, r) => s + Number(r['Gross P/L'] ?? 0), 0) >= 0 ? '+' : ''}
                          {openRows.reduce((s, r) => s + Number(r['Gross P/L'] ?? 0), 0).toFixed(2)} EUR
                        </div>
                      </div>
                      <div className="card-sep" />
                      <div style={{ marginTop: 14 }}>
                        <div style={{ fontSize: 10, color: 'var(--text-3)', fontWeight: 600, letterSpacing: '0.08em', textTransform: 'uppercase', marginBottom: 4 }}>Current value</div>
                        <div className="card-value" style={{ fontSize: 18, marginBottom: 0, textAlign: 'right' }}>{openRows.reduce((s, r) => s + Number(r['Volume'] ?? 0) * Number(r['Market price'] ?? 0), 0).toFixed(2)} EUR</div>
                      </div>
                    </div>

                  {/* Symbol summary — tax exempt volume */}
                  {(() => {
                    const symbolMap = new Map<string, { total: number; exempt: number }>()
                    for (const r of openRows) {
                      const sym = String(r['Symbol'] ?? '—')
                      const vol = Number(r['Volume'] ?? 0)
                      const exempt = r['Tax exempt'] === 'Yes'
                      if (!symbolMap.has(sym)) symbolMap.set(sym, { total: 0, exempt: 0 })
                      const e = symbolMap.get(sym)!
                      e.total += vol
                      if (exempt) e.exempt += vol
                    }
                    const symbols = [...symbolMap.entries()].sort((a, b) => b[1].total - a[1].total)
                    return (
                      <div className="table-shell appear-1" style={{ marginBottom: 24 }}>
                        <div className="table-topbar">
                          <div className="table-topbar-left">
                            <span className="table-label">By Symbol</span>
                            <span className="table-count">{symbols.length}</span>
                          </div>
                        </div>
                        <table>
                          <thead>
                            <tr>
                              <th style={{ textAlign: 'left' }}>Symbol</th>
                              <th>Total Volume</th>
                              <th>Not Exempt Volume</th>
                              <th>Exempt Volume</th>
                            </tr>
                          </thead>
                          <tbody>
                            {symbols.map(([sym, { total, exempt }]) => {
                              const notExempt = total - exempt
                              const allExempt = notExempt < 0.0001
                              return (
                                <tr key={sym}>
                                  <td style={{ textAlign: 'left', fontFamily: 'Syne, sans-serif', fontWeight: 600, color: 'var(--text-1)' }}>{sym}</td>
                                  <td>{total.toFixed(4)}</td>
                                  <td className={notExempt > 0.0001 ? 'num-red' : 'pnl-zero'}>{notExempt > 0.0001 ? notExempt.toFixed(4) : '—'}</td>
                                  <td className="num-green">{exempt > 0 ? exempt.toFixed(4) : '—'}</td>

                                </tr>
                              )
                            })}
                          </tbody>
                        </table>
                      </div>
                    )
                  })()}

                  {/* ── What can I sell? dashboard ── */}
                  {todayFx > 0 && (() => {
                    const remaining = Math.max(0, taxablePurchaseCZK - taxableSaleCZK)
                    // Build per-symbol summary from openRows
                    const symMap = new Map<string, { volume: number; marketPrice: number }>()
                    for (const r of openRows) {
                      const sym = String(r['Symbol'] ?? '—')
                      const vol = Number(r['Volume'] ?? 0)
                      const mp  = Number(r['Market price'] ?? 0)
                      if (!symMap.has(sym)) symMap.set(sym, { volume: 0, marketPrice: mp })
                      symMap.get(sym)!.volume += vol
                    }
                    const symbols = [...symMap.entries()].map(([sym, { volume, marketPrice }]) => {
                      const estFullCZK = volume * marketPrice * todayFx
                      const maxSellCZK  = Math.min(Math.max(remaining, 0), estFullCZK)
                      const maxSellVol  = todayFx && marketPrice ? maxSellCZK / (marketPrice * todayFx) : 0
                      const canSellAll  = estFullCZK <= remaining
                      return { sym, volume, marketPrice, maxSellCZK, maxSellVol, canSellAll }
                    }).sort((a, b) => b.maxSellVol - a.maxSellVol)

                    return (
                      <div className="table-shell appear-2" style={{ marginBottom: 24 }}>
                        <div className="table-topbar">
                          <div className="table-topbar-left">
                            <span className="table-label">What can I sell?</span>
                            <span style={{ fontFamily: 'var(--mono)', fontSize: 11, color: remaining > 0 ? 'var(--green)' : 'var(--red)', background: remaining > 0 ? 'rgba(61,224,176,0.08)' : 'rgba(255,128,144,0.08)', border: `1px solid ${remaining > 0 ? 'rgba(61,224,176,0.2)' : 'rgba(255,128,144,0.2)'}`, borderRadius: 4, padding: '2px 8px' }}>
                              {remaining > 0 ? `${Math.round(remaining).toLocaleString('cs-CZ')} Kč available` : `No budget remaining`}
                            </span>
                          </div>
                          {todayFx > 0 && (
                            <span style={{ fontFamily: 'var(--mono)', fontSize: 11, color: 'var(--text-3)' }}>
                              EUR/CZK {todayFx.toFixed(3)} <span style={{ fontSize: 10 }}>ČNB today</span>
                            </span>
                          )}
                        </div>
                        <table>
                          <thead>
                            <tr>
                              <th style={{ textAlign: 'left' }}>Symbol</th>
                              <th>Total Volume</th>
                              <th>Market Price</th>
                              <th>Max Sellable Volume</th>
                              <th>Max Sellable CZK</th>
                            </tr>
                          </thead>
                          <tbody>
                            {symbols.map(({ sym, volume, marketPrice, maxSellCZK, maxSellVol, canSellAll }) => (
                              <tr key={sym}>
                                <td style={{ textAlign: 'left', fontFamily: 'Syne, sans-serif', fontWeight: 600, color: 'var(--text-1)' }}>{sym}</td>
                                <td>{volume.toFixed(4)}</td>
                                <td>{marketPrice.toFixed(4)}</td>
                                <td className={canSellAll ? 'num-green' : 'num-red'}>{maxSellVol.toFixed(4)}</td>
                                <td className={canSellAll ? 'num-green' : 'num-red'}>{Math.round(maxSellCZK).toLocaleString('cs-CZ')} Kč</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )
                  })()}
                  </div>

                  {/* Open positions table */}
                  <div className="table-shell appear-3">
                    <div className="table-topbar">
                      <div className="table-topbar-left">
                        <span className="table-label">Open Positions</span>
                        <span className="table-count">{openRows.length}</span>
                      </div>
                    </div>
                    <div className="table-scroll" style={{ overflowX: 'auto' }}>
                      <table>
                        <thead>
                          <tr>
                            {[['Symbol','left'],['Type','right'],['Volume','right'],['Open Date','right'],['Days Held','right'],['Open Price','right'],['Market Price','right'],['Purchase EUR','right'],['Gross P/L','right'],['Tax Exempt','center'],['Days to Exempt','center']].map(([col, align]) => (
                              <th
                                key={col}
                                style={{ textAlign: align as any }}
                                className={openSortKey === col ? 'th-active' : ''}
                                onClick={() => {
                                  if (openSortKey === col) setOpenSortDir(p => p === 'asc' ? 'desc' : 'asc')
                                  else { setOpenSortKey(col); setOpenSortDir('asc') }
                                }}
                              >
                                {col}{openSortKey && openSortKey === col ? (openSortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {sortedOpenRows.map((r, i) => {
                            const pnl = Number(r['Gross P/L'] ?? 0)
                            const exempt = r['Tax exempt'] === 'Yes'
                            const daysToExempt = Number(r['Days to exempt'] ?? 0)
                            return (
                              <tr key={i}>
                                <td style={{ textAlign: 'left', fontFamily: 'Syne, sans-serif', fontWeight: 600, color: 'var(--text-1)' }}>{String(r['Symbol'] ?? '')}</td>
                                <td>{String(r['Type'] ?? '')}</td>
                                <td>{String(r['Volume'] ?? '')}</td>
                                <td>{String(r['_openDisplay'] ?? '').split(' ')[0]}</td>
                                <td>{String(r['Days held'] ?? '')}</td>
                                <td>{String(r['Open price'] ?? '')}</td>
                                <td>{String(r['Market price'] ?? '')}</td>
                                <td>{String(r['Purchase value'] ?? '')}</td>
                                <td style={{ color: pnl >= 0 ? 'var(--green)' : 'var(--red)' }}>{pnl >= 0 ? '+' : ''}{pnl.toFixed(2)}</td>
                                <td style={{ color: exempt ? 'var(--green)' : 'var(--red)', textAlign: 'center' }}>{r['Tax exempt'] ?? ''}</td>
                                <td style={{ color: daysToExempt <= 180 && !exempt ? 'var(--amber)' : 'var(--text-3)', textAlign: 'center' }}>{exempt ? '—' : daysToExempt}</td>
                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              )}
            </div>
          )}

          {/* ── Trades table — shown in overview and closed views ── */}
          {activeView === 'overview' && (
            <div className="table-shell appear-4">
              <div className="table-topbar">
                <div className="table-topbar-left">
                  <span className="table-label">Closed Positions</span>
                  <span className="table-count">{filteredRows.length}</span>
                </div>
                <button className="export-btn" onClick={exportToExcel}>
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                  </svg>
                  Export Excel
                </button>
              </div>
              <div className="table-scroll" style={{ overflowX: 'auto' }}>
                <table>
                  <thead>
                    <tr>
                      {visibleKeys.map(k => (
                        <th key={k} className={sortKey === k ? 'th-active' : ''} onClick={() => { if (sortKey === k) setSortDir(p => p === 'asc' ? 'desc' : 'asc'); else { setSortKey(k); setSortDir('asc') } }}>
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
                            {k === 'Open time'  ? String(r['_openDisplay']  ?? '').split(' ')[0]
                            : k === 'Close time' ? String(r['_closeDisplay'] ?? '').split(' ')[0]
                            : String(r[k] ?? '')}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          </div>
        )}

      </div>{/* end app-shell */}
    </div>
  )
}
