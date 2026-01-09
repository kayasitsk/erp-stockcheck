import React, { useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'

const SIZES = ['S', 'M', 'L', 'XL', 'XXL', '3XL', '4XL']
const SIZE_MATCH_ORDER = ['4XL','3XL','XXL','XL','L','M','S'] // longest-first

const I18N = {
  th: {
    appTitle: '{t.appTitle}',
    appSub: '{t.appSub}',
    runsLocal: '{t.runsLocal}',
    step1: '{t.step1}',
    dropPrimary: '{t.dropPrimary}',
    dropSecondary: '{t.dropSecondary}',
    pickFile: '{t.pickFile}',
    clear: '{t.clear}',
    updatedDate: '{t.updatedDate}',
    filterModel: '{t.filterModel}',
    all: '{t.all}',
    search: '{t.search}',
    searchPlaceholder: '{t.searchPlaceholder}',
    generate: '{t.generate}',
    processing: '{t.processing}',
    note: '{t.note}',
    model: '{t.model}',
    color: '{t.color}',
    emptyStateNoData: '{t.emptyStateNoData}',
    emptyStateLoading: '{t.emptyStateLoading}',
    emptyStateNoMatch: '{t.emptyStateNoMatch}',
    step2: '{t.step2}',
    uploadedFiles: '{t.uploadedFiles}',
    readableRows: '{t.readableRows}',
    rowCount: 'จำนวนแถว ({t.model}+{t.color})',
    badSku: '{t.badSku}',
    dlError: '{t.dlError}',
    sizes: '{t.sizes}',
    missing: '{t.missing}',
    aggregation: '{t.aggregation}',
    missingZero: '{t.missingZero}',
    aggSum: '{t.aggSum}',
    foot: '{t.foot}',
    lang: 'ภาษา',
    langTH: 'ไทย',
    langZH: '中文',
    checksTitle: 'ความถูกต้องของตัวเลข',
    checkRaw: 'ผลรวมสต็อก (จากไฟล์หลัง parse)',
    checkMatrix: 'ผลรวมสต็อก (หลังจัดตาราง)',
    checkDiff: 'ส่วนต่าง',
    checkOk: 'ตรงกัน',
    checkWarn: 'ไม่ตรงกัน (โปรดตรวจสอบ)',
    perFile: 'แยกตามไฟล์',
  },
  zh: {
    appTitle: 'ERP → 库存盘点表（一次合并 T009/T111）',
    appSub: '支持同时上传多个 ERP 文件 → 自动把 SKU 拆分为 型号/颜色/尺码，并生成与模板一致的盘点表（缺失填 0）',
    runsLocal: '本地运行 • React + Vite • ExcelJS',
    step1: '1) 上传 ERP 文件',
    dropPrimary: '把 .xlsx 文件拖到这里，或点击选择文件',
    dropSecondary: '支持多文件（如 T009.xlsx + T111.xlsx）。即使不同型号颜色同名也不会混淆，因为键是（型号, 颜色, 尺码）',
    pickFile: '选择文件',
    clear: '清空',
    updatedDate: '更新日期',
    filterModel: '筛选型号',
    all: '全部',
    search: '搜索',
    searchPlaceholder: '例如 T009 或 darkgreen',
    generate: '生成 Excel',
    processing: '处理中…',
    note: '说明：程序会先按表头名称识别列；若识别失败，会回退到 B 列（SKU）与 I 列（可售库存）',
    model: '型号',
    color: '颜色',
    emptyStateNoData: '暂无数据 — 请先上传 ERP 文件',
    emptyStateLoading: '正在读取文件…',
    emptyStateNoMatch: '没有匹配到数据',
    step2: '2) 导出前检查',
    uploadedFiles: '已上传文件',
    readableRows: '读取到的 SKU 行数',
    rowCount: '行数（型号+颜色）',
    badSku: '无法解析的 SKU',
    dlError: '下载错误报告（CSV）',
    sizes: '尺码',
    missing: '缺失填充',
    aggregation: '汇总方式',
    missingZero: '填 0',
    aggSum: '求和（避免多文件重复丢失）',
    foot: '如有 SKU 格式异常（例如缺少 - 或末尾没有尺码），请下载 CSV 错误报告核对后再导出',
    lang: '语言',
    langTH: 'ไทย',
    langZH: '中文',
    checksTitle: '数字准确性校验',
    checkRaw: '库存合计（解析后）',
    checkMatrix: '库存合计（生成表后）',
    checkDiff: '差值',
    checkOk: '一致',
    checkWarn: '不一致（请核对）',
    perFile: '按文件查看',
  }
}

function formatBytes(bytes){
  if (!Number.isFinite(bytes)) return ''
  const units = ['B','KB','MB','GB']
  let b = bytes
  let i = 0
  while (b >= 1024 && i < units.length - 1) { b/=1024; i++ }
  return `${b.toFixed(i===0?0:1)} ${units[i]}`
}

function todayISO(){
  const d = new Date()
  const pad = n => String(n).padStart(2,'0')
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`
}

function parseSku(skuRaw){
  const sku = String(skuRaw ?? '').trim()
  if (!sku) return { ok:false, reason:'empty' }
  const dashIdx = sku.indexOf('-')
  if (dashIdx < 1) return { ok:false, reason:'missing-dash' }
  const model = sku.slice(0, dashIdx).trim()
  const tail = sku.slice(dashIdx+1).trim()
  if (!model || !tail) return { ok:false, reason:'bad-format' }

  // match size suffix
  const upperTail = tail.toUpperCase()
  let size = null
  for (const cand of SIZE_MATCH_ORDER){
    if (upperTail.endsWith(cand)){
      size = cand
      break
    }
  }
  if (!size) return { ok:false, reason:'unknown-size' }

  const color = tail.slice(0, tail.length - size.length).trim()
  if (!color) return { ok:false, reason:'missing-color' }

  return { ok:true, model, color, size, sku }
}

function tryFindColumnIndex(headers, predicates){
  // headers: array of strings
  for (let i=0; i<headers.length; i++){
    const h = (headers[i] ?? '').toString().trim()
    if (!h) continue
    const ok = predicates.some(p => p(h))
    if (ok) return i
  }
  return -1
}

function normalizeHeaderRow(rowValues){
  // ExcelJS row.values is 1-based; rowValues may be sparse
  const arr = []
  for (let c=1; c<rowValues.length; c++){
    const v = rowValues[c]
    arr.push((v ?? '').toString().trim())
  }
  return arr
}

async function readErpFiles(files){
  const rows = []
  const errors = []
  let totalRead = 0

  for (const file of files){
    const arrayBuffer = await file.arrayBuffer()
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(arrayBuffer)

    // choose first worksheet by default
    const ws = wb.worksheets[0]
    if (!ws){
      errors.push({ file: file.name, sku: '', reason: 'no-worksheet' })
      continue
    }

    // header row assumed at row 1 in your ERP export (based on screenshot)
    // but we still attempt detection: pick first row with >=5 non-empty cells
    let headerRowNumber = 1
    for (let r=1; r<=Math.min(ws.rowCount, 10); r++){
      const rv = ws.getRow(r).values
      const headers = normalizeHeaderRow(rv)
      const nonEmpty = headers.filter(x => x).length
      if (nonEmpty >= 5){
        headerRowNumber = r
        break
      }
    }

    const headerValues = normalizeHeaderRow(ws.getRow(headerRowNumber).values)

    // Attempt to find SKU + stock-ready columns by header name; fallback to B and I
    const skuCol = (() => {
      const idx = tryFindColumnIndex(headerValues, [
        h => h.toUpperCase() === 'SKU',
        h => h.toUpperCase().includes('SKU'),
        h => h.toUpperCase().includes('สินค้าคงคลัง') && h.toUpperCase().includes('SKU'),
      ])
      return idx >= 0 ? idx + 1 : 2 // ExcelJS col is 1-based; B=2
    })()

    const stockCol = (() => {
      const idx = tryFindColumnIndex(headerValues, [
        h => h.includes('สต็อกพร้อมขาย'),
        h => h.includes('พร้อมขาย'),
        h => h.toUpperCase().includes('AVAILABLE'),
        h => h.toUpperCase().includes('READY'),
      ])
      return idx >= 0 ? idx + 1 : 9 // I=9
    })()

    // iterate rows below header
    for (let r=headerRowNumber+1; r<=ws.rowCount; r++){
      const row = ws.getRow(r)
      const skuVal = row.getCell(skuCol).value
      const stockVal = row.getCell(stockCol).value

      const sku = (skuVal?.text ?? skuVal ?? '').toString().trim()
      if (!sku) continue

      // parse stock number safely
      let stock = 0
      if (typeof stockVal === 'number') stock = stockVal
      else if (typeof stockVal === 'string') stock = Number(stockVal.replace(/,/g,''))
      else if (stockVal && typeof stockVal === 'object' && 'result' in stockVal) stock = Number(stockVal.result)
      else if (stockVal?.text) stock = Number(String(stockVal.text).replace(/,/g,''))
      if (!Number.isFinite(stock)) stock = 0
      stock = Math.trunc(stock)

      const parsed = parseSku(sku)
      if (!parsed.ok){
        errors.push({ file: file.name, sku, reason: parsed.reason })
        continue
      }

      rows.push({
        file: file.name,
        sku: parsed.sku,
        model: parsed.model,
        color: parsed.color,
        size: parsed.size,
        stock,
      })
      totalRead++
    }
  }

  return { rows, errors, totalRead }
}

function buildMatrix(rows){
  // key: model||color
  const map = new Map()
  const models = new Set()

  for (const r of rows){
    models.add(r.model)
    const key = `${r.model}||${r.color}`
    if (!map.has(key)){
      const init = { model: r.model, color: r.color }
      for (const s of SIZES) init[s] = 0
      map.set(key, init)
    }
    const obj = map.get(key)
    if (SIZES.includes(r.size)){
      obj[r.size] = (obj[r.size] ?? 0) + (r.stock ?? 0) // sum to be safe
    }
  }

  const out = Array.from(map.values())
  out.sort((a,b) => (a.model.localeCompare(b.model) || a.color.localeCompare(b.color)))
  return { matrix: out, models: Array.from(models).sort((a,b)=>a.localeCompare(b)) }
}

async function generateExcel({ matrix, updatedISO }){
  const wb = new ExcelJS.Workbook()
  wb.creator = 'ERP Stock Checker'
  wb.created = new Date()

  const ws = wb.addWorksheet('ใบเช็คสต็อก', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 3, xSplit: 0 }]
  })

  // Column widths (approx like your form)
  ws.getColumn(1).width = 10 // {t.model}
  ws.getColumn(2).width = 16 // {t.color}
  const sizeColsStart = 3
  for (let i=0; i<SIZES.length; i++){
    ws.getColumn(sizeColsStart+i).width = 9
  }

  // Title row (row 2 in your screenshot, but we'll use row 1-2 area)
  ws.mergeCells('C1','G2')
  const titleCell = ws.getCell('C1')
  titleCell.value = 'ใบเช็คสต็อก'
  titleCell.font = { size: 28, bold: true }
  titleCell.alignment = { vertical:'middle', horizontal:'center' }

  // Update date box on the right (H1:I2)
  ws.mergeCells('H1','I1')
  ws.getCell('H1').value = '{t.updatedDate}'
  ws.getCell('H1').font = { bold: true, size: 14 }
  ws.getCell('H1').alignment = { vertical:'middle', horizontal:'center' }

  ws.mergeCells('H2','I2')
  // show as dd/mm/yyyy
  const d = new Date(updatedISO + 'T00:00:00')
  const dd = String(d.getDate()).padStart(2,'0')
  const mm = String(d.getMonth()+1).padStart(2,'0')
  const yyyy = d.getFullYear()
  ws.getCell('H2').value = `${dd}/${mm}/${yyyy}`
  ws.getCell('H2').alignment = { vertical:'middle', horizontal:'center' }
  ws.getCell('H2').font = { bold: true, size: 14 }

  // Header row (Row 3)
  const headerRowNum = 3
  ws.getCell(`A${headerRowNum}`).value = '{t.model}'
  ws.getCell(`B${headerRowNum}`).value = '{t.color}'
  for (let i=0; i<SIZES.length; i++){
    ws.getCell(headerRowNum, sizeColsStart+i).value = SIZES[i]
  }

  // Styling helpers
  const thin = { style:'thin', color:{ argb:'FF000000' } }
  const thick = { style:'medium', color:{ argb:'FF000000' } }

  function setBorder(cell, which='thin'){
    const b = which==='thick' ? thick : thin
    cell.border = { top:b, left:b, bottom:b, right:b }
  }

  // Apply header styles
  for (let c=1; c<=2+SIZES.length; c++){
    const cell = ws.getCell(headerRowNum, c)
    cell.font = { bold: true }
    cell.alignment = { vertical:'middle', horizontal:'center' }
    setBorder(cell, 'thick')
  }

  // Data rows start at row 4
  let r = headerRowNum + 1
  for (const item of matrix){
    ws.getCell(r,1).value = item.model
    ws.getCell(r,2).value = item.color

    ws.getCell(r,1).alignment = { vertical:'middle', horizontal:'left' }
    ws.getCell(r,2).alignment = { vertical:'middle', horizontal:'left' }

    setBorder(ws.getCell(r,1))
    setBorder(ws.getCell(r,2))

    for (let i=0; i<SIZES.length; i++){
      const v = Number(item[SIZES[i]] ?? 0)
      const cell = ws.getCell(r, sizeColsStart+i)
      cell.value = v
      cell.alignment = { vertical:'middle', horizontal:'center' }
      setBorder(cell)

      // subtle conditional formatting-ish: color low/zero via font
      if (v === 0){
        cell.font = { color: { argb:'FF6B7280' } } // grey
      } else if (v <= 5){
        cell.font = { color: { argb:'FFB45309' }, bold: true } // amber
      } else {
        cell.font = { color: { argb:'FF111827' } } // near black
      }
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFFF' } }
    }
    r++
  }

  // Surround the "update date" box with border
  for (const addr of ['H1','I1','H2','I2']){
    setBorder(ws.getCell(addr), 'thick')
  }
  ws.getCell('H1').fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF3F4F6' } }
  ws.getCell('H2').fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFFF' } }
  ws.getCell('C1').fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFFF' } }

  // Add outer border around the main table area (A3:I lastRow)
  const lastRow = Math.max(headerRowNum+1, r-1)
  const lastCol = 2 + SIZES.length // I
  for (let rr=headerRowNum; rr<=lastRow; rr++){
    for (let cc=1; cc<=lastCol; cc++){
      const cell = ws.getCell(rr, cc)
      // overwrite borders on edges to thick
      const top = (rr === headerRowNum) ? thick : (cell.border?.top ?? thin)
      const bottom = (rr === lastRow) ? thick : (cell.border?.bottom ?? thin)
      const left = (cc === 1) ? thick : (cell.border?.left ?? thin)
      const right = (cc === lastCol) ? thick : (cell.border?.right ?? thin)
      cell.border = { top, bottom, left, right }
      // fill white for printable look
      if (!cell.fill){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFFF' } }
      }
    }
  }

  // Make worksheet background printable white by setting fills in used range (done above)
  // Return as Blob
  const buffer = await wb.xlsx.writeBuffer()
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

function downloadBlob(blob, filename){
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}

export default function App(){
  const inputRef = useRef(null)
  const [lang, setLang] = useState('th')
  const t = I18N[lang]

  const [files, setFiles] = useState([])
  const [updatedISO, setUpdatedISO] = useState(todayISO())
  const [loading, setLoading] = useState(false)

  const [rows, setRows] = useState([])
  const [errors, setErrors] = useState([])

  const [modelFilter, setModelFilter] = useState('ALL')
  const [query, setQuery] = useState('')

  const { matrix, models } = useMemo(() => buildMatrix(rows), [rows])

  const filteredMatrix = useMemo(() => {
    const q = query.trim().toLowerCase()
    return matrix.filter(item => {
      if (modelFilter !== 'ALL' && item.model !== modelFilter) return false
      if (!q) return true
      if (item.model.toLowerCase().includes(q)) return true
      if (item.color.toLowerCase().includes(q)) return true
      return false
    })
  }, [matrix, modelFilter, query])

  const stats = useMemo(() => {
    const totalSkuRows = rows.length
    const badSku = errors.length
    const uniquePairs = new Set(rows.map(r => `${r.model}||${r.color}||${r.size}`)).size
    const uniqueRows = new Set(rows.map(r => `${r.model}||${r.color}`)).size
    return { totalSkuRows, badSku, uniquePairs, uniqueRows }
  }, [rows, errors])

const checks = useMemo(() => {
  const rawTotal = rows.reduce((acc, r) => acc + (Number(r.stock) || 0), 0)

  const matrixTotal = matrix.reduce((acc, item) => {
    let s = 0
    for (const size of SIZES) s += (Number(item[size]) || 0)
    return acc + s
  }, 0)

  const perFile = new Map()
  for (const r of rows){
    const key = r.file || 'unknown'
    perFile.set(key, (perFile.get(key) || 0) + (Number(r.stock) || 0))
  }

  const diff = rawTotal - matrixTotal
  return {
    rawTotal,
    matrixTotal,
    diff,
    ok: diff === 0,
    perFile: Array.from(perFile.entries()).sort((a,b)=>a[0].localeCompare(b[0])),
  }
}, [rows, matrix])

  async function handleFilesSelected(fileList){
    const arr = Array.from(fileList || []).filter(f => f.name.toLowerCase().endsWith('.xlsx') || f.name.toLowerCase().endsWith('.xls'))
    setFiles(arr)
    setRows([])
    setErrors([])
    setModelFilter('ALL')
    setQuery('')

    if (arr.length === 0) return

    setLoading(true)
    try{
      const { rows, errors } = await readErpFiles(arr)
      setRows(rows)
      setErrors(errors)
    } finally {
      setLoading(false)
    }
  }

  function onDrop(e){
    e.preventDefault()
    const dropped = e.dataTransfer.files
    handleFilesSelected(dropped)
  }

  function onDragOver(e){
    e.preventDefault()
  }

  async function onGenerate(){
    setLoading(true)
    try{
      const blob = await generateExcel({ matrix: filteredMatrix.length ? filteredMatrix : matrix, updatedISO })
      const d = updatedISO.replaceAll('-','')
      const filename = `ใบเช็คสต็อก_${d}.xlsx`
      downloadBlob(blob, filename)
    } finally {
      setLoading(false)
    }
  }

  async function onDownloadErrors(){
    // simple CSV for errors
    const header = 'file,sku,reason\n'
    const lines = errors.map(e => {
      const safe = (s) => `"${String(s??'').replaceAll('"','""')}"`
      return [safe(e.file), safe(e.sku), safe(e.reason)].join(',')
    }).join('\n')
    const blob = new Blob([header + lines], { type:'text/csv;charset=utf-8' })
    downloadBlob(blob, `errors_${updatedISO.replaceAll('-','')}.csv`)
  }

  const canExport = files.length > 0 && rows.length > 0 && !loading

  return (
    <div className="container">
      <div className="header">
        <div>
          <div className="hTitle">{t.appTitle}</div>
          <div className="hSub">
            {t.appSub}
          </div>
        </div>
        <div style={{display:'flex', flexDirection:'column', gap:10, alignItems:'flex-end'}}>
  <div className="badge">{t.runsLocal}</div>
  <div className="input" style={{padding:'8px 10px'}}>
    <label>{t.lang}</label>
    <select value={lang} onChange={(e)=>setLang(e.target.value)}>
      <option value="th">{t.langTH}</option>
      <option value="zh">{t.langZH}</option>
    </select>
  </div>
</div>
      </div>

      <div className="grid">
        <div className="card">
          <div className="cardTitle">{t.step1}</div>

          <div className="drop" onDrop={onDrop} onDragOver={onDragOver}>
            <div className="dropLeft">
              <div className="dropPrimary">{t.dropPrimary}</div>
              <div className="dropSecondary">{t.dropSecondary}</div>
            </div>
            <div style={{display:'flex', gap:8}}>
              <button className="btn secondary" onClick={() => { inputRef.current?.click() }} disabled={loading}>{t.pickFile}</button>
              <button className="btn" onClick={() => handleFilesSelected([])} disabled={loading || (files.length===0 && rows.length===0 && errors.length===0)}>{t.clear}</button>
            </div>
          </div>

          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            multiple
            style={{display:'none'}}
            onChange={(e) => handleFilesSelected(e.target.files)}
          />

          {files.length > 0 && (
            <div className="fileList">
              {files.map(f => (
                <div className="fileItem" key={f.name}>
                  <div className="fileName">{f.name}</div>
                  <div className="fileMeta">{formatBytes(f.size)}</div>
                </div>
              ))}
            </div>
          )}

          <div className="row">
            <div className="input">
              <label>{t.updatedDate}</label>
              <input type="date" value={updatedISO} onChange={(e)=>setUpdatedISO(e.target.value)} />
            </div>

            <div className="input">
              <label>{t.filterModel}</label>
              <select value={modelFilter} onChange={(e)=>setModelFilter(e.target.value)}>
                <option value="ALL">{t.all}</option>
                {models.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
            </div>

            <div className="input" style={{flex:1, minWidth: 220}}>
              <label>{t.search}</label>
              <input placeholder={t.searchPlaceholder} value={query} onChange={(e)=>setQuery(e.target.value)} />
            </div>

            <button className="btn" disabled={!canExport} onClick={onGenerate}>
              {loading ? '{t.processing}' : '{t.generate}'}
            </button>
          </div>

          <div className="footerHelp">
            {t.note}
          </div>

          <div className="tableWrap">
            <table>
              <thead>
                <tr>
                  <th className="mono">{t.model}</th>
                  <th className="mono">{t.color}</th>
                  {SIZES.map(s => <th className="mono" key={s}>{s}</th>)}
                </tr>
              </thead>
              <tbody>
                {(filteredMatrix.slice(0, 30)).map((row, idx) => (
                  <tr key={`${row.model}||${row.color}||${idx}`}>
                    <td className="mono">{row.model}</td>
                    <td className="mono">{row.color}</td>
                    {SIZES.map(s => <td className="mono" key={s}>{row[s]}</td>)}
                  </tr>
                ))}
                {filteredMatrix.length === 0 && (
                  <tr>
                    <td colSpan={2+SIZES.length} style={{color:'rgba(255,255,255,0.65)'}}>
                      {files.length===0 ? t.emptyStateNoData : (loading ? t.emptyStateLoading : t.emptyStateNoMatch)}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

        </div>

        <div className="card">
          <div className="cardTitle">{t.step2}</div>
          <div className="statRow">
            <div className="stat">
              <div className="left">
                <div className={`dot ${files.length>0 ? 'good' : ''}`}></div>
                <div className="statLabel">{t.uploadedFiles}</div>
              </div>
              <div className="statValue">{files.length}</div>
            </div>

            <div className="stat">
              <div className="left">
                <div className={`dot ${rows.length>0 ? 'good' : (files.length>0 ? 'warn' : '')}`}></div>
                <div className="statLabel">{t.readableRows}</div>
              </div>
              <div className="statValue">{stats.totalSkuRows}</div>
            </div>

            <div className="stat">
              <div className="left">
                <div className={`dot ${stats.uniqueRows>0 ? 'good' : ''}`}></div>
                <div className="statLabel">จำนวนแถว ({t.model}+{t.color})</div>
              </div>
              <div className="statValue">{stats.uniqueRows}</div>
            </div>

            <div className="stat">
              <div className="left">
                <div className={`dot ${stats.badSku===0 ? 'good' : 'warn'}`}></div>
                <div className="statLabel">{t.badSku}</div>
              </div>
              <div className="statValue">{stats.badSku}</div>
            </div>

            {errors.length > 0 && (
              <button className="btn secondary" onClick={onDownloadErrors} disabled={loading}>
                {t.dlError}
              </button>
            )}

            <div className="kv">
              <div className="pill"><strong>{t.sizes}:</strong> {SIZES.join(', ')}</div>
              <div className="pill"><strong>{t.missing}:</strong> {t.missingZero}</div>
              <div className="pill"><strong>{t.aggregation}:</strong> {t.aggSum}</div>
            </div>

<div style={{marginTop: 12}}>
  <div className="cardTitle" style={{marginBottom: 10}}>{t.checksTitle}</div>
  <div className="statRow">
    <div className="stat">
      <div className="left">
        <div className={`dot ${checks.ok ? 'good' : (rows.length>0 ? 'warn' : '')}`}></div>
        <div className="statLabel">{t.checkRaw}</div>
      </div>
      <div className="statValue">{checks.rawTotal}</div>
    </div>

    <div className="stat">
      <div className="left">
        <div className={`dot ${checks.ok ? 'good' : (rows.length>0 ? 'warn' : '')}`}></div>
        <div className="statLabel">{t.checkMatrix}</div>
      </div>
      <div className="statValue">{checks.matrixTotal}</div>
    </div>

    <div className="stat">
      <div className="left">
        <div className={`dot ${checks.ok ? 'good' : (rows.length>0 ? 'warn' : '')}`}></div>
        <div className="statLabel">{t.checkDiff}</div>
      </div>
      <div className="statValue">{checks.diff}</div>
    </div>

    {rows.length > 0 && (
      <div className="footerHelp" style={{marginTop: 6}}>
        {checks.ok ? <span style={{color:'rgba(34,197,94,0.95)'}}>{t.checkOk}</span> : <span style={{color:'rgba(245,158,11,0.95)'}}>{t.checkWarn}</span>}
      </div>
    )}

    {checks.perFile.length > 0 && (
      <div className="tableWrap" style={{marginTop: 10}}>
        <table>
          <thead>
            <tr>
              <th className="mono">{t.perFile}</th>
              <th className="mono">{t.checkRaw}</th>
            </tr>
          </thead>
          <tbody>
            {checks.perFile.map(([name, sum]) => (
              <tr key={name}>
                <td className="mono">{name}</td>
                <td className="mono">{sum}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )}
  </div>
</div>
          </div>

          <div className="footerHelp">
            {t.foot}
          </div>
        </div>
      </div>
    </div>
  )
}
