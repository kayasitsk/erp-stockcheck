import React, { useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'

const SIZES = ['S', 'M', 'L', 'XL', 'XXL', '3XL', '4XL']
const SIZE_MATCH_ORDER = ['4XL','3XL','XXL','XL','L','M','S'] // longest-first

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
  ws.getColumn(1).width = 10 // รุ่น
  ws.getColumn(2).width = 16 // สี
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
  ws.getCell('H1').value = 'อัปเดตวันที่'
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
  ws.getCell(`A${headerRowNum}`).value = 'รุ่น'
  ws.getCell(`B${headerRowNum}`).value = 'สี'
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
          <div className="hTitle">ERP → ใบเช็คสต็อก (T009/T111 รวมรอบเดียว)</div>
          <div className="hSub">
            อัปโหลดไฟล์ ERP ได้หลายไฟล์พร้อมกัน → โปรแกรมจะแยก SKU เป็น รุ่น/สี/ไซส์ และทำตารางเช็คสต็อกแบบในฟอร์มให้ทันที (ช่องที่ไม่มีให้เป็น 0)
          </div>
        </div>
        <div className="badge">Runs locally • React + Vite • ExcelJS</div>
      </div>

      <div className="grid">
        <div className="card">
          <div className="cardTitle">1) อัปโหลดไฟล์ ERP</div>

          <div className="drop" onDrop={onDrop} onDragOver={onDragOver}>
            <div className="dropLeft">
              <div className="dropPrimary">ลากไฟล์ .xlsx มาวางที่นี่ หรือกดเลือกไฟล์</div>
              <div className="dropSecondary">รองรับหลายไฟล์พร้อมกัน (เช่น T009.xlsx + T111.xlsx) และจะไม่สับสนแม้สีซ้ำ เพราะคีย์เป็น (รุ่น, สี, ไซส์)</div>
            </div>
            <div style={{display:'flex', gap:8}}>
              <button className="btn secondary" onClick={() => { inputRef.current?.click() }} disabled={loading}>เลือกไฟล์</button>
              <button className="btn" onClick={() => handleFilesSelected([])} disabled={loading || (files.length===0 && rows.length===0 && errors.length===0)}>ล้าง</button>
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
              <label>อัปเดตวันที่</label>
              <input type="date" value={updatedISO} onChange={(e)=>setUpdatedISO(e.target.value)} />
            </div>

            <div className="input">
              <label>Filter รุ่น</label>
              <select value={modelFilter} onChange={(e)=>setModelFilter(e.target.value)}>
                <option value="ALL">ทั้งหมด</option>
                {models.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
            </div>

            <div className="input" style={{flex:1, minWidth: 220}}>
              <label>ค้นหา</label>
              <input placeholder="เช่น T009 หรือ darkgreen" value={query} onChange={(e)=>setQuery(e.target.value)} />
            </div>

            <button className="btn" disabled={!canExport} onClick={onGenerate}>
              {loading ? 'กำลังประมวลผล…' : 'Generate Excel'}
            </button>
          </div>

          <div className="footerHelp">
            หมายเหตุ: โปรแกรมพยายามหา column จากชื่อหัวตารางก่อน ถ้าไม่เจอจะ fallback ไปที่คอลัมน์ B (SKU) และ I (สต็อกพร้อมขาย) ตามไฟล์ ERP ของคุณ
          </div>

          <div className="tableWrap">
            <table>
              <thead>
                <tr>
                  <th className="mono">รุ่น</th>
                  <th className="mono">สี</th>
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
                      {files.length===0 ? 'ยังไม่มีข้อมูล — อัปโหลดไฟล์ ERP ก่อน' : (loading ? 'กำลังอ่านไฟล์…' : 'ไม่พบข้อมูลตามตัวกรอง')}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

        </div>

        <div className="card">
          <div className="cardTitle">2) ตรวจสอบก่อน Export</div>
          <div className="statRow">
            <div className="stat">
              <div className="left">
                <div className={`dot ${files.length>0 ? 'good' : ''}`}></div>
                <div className="statLabel">ไฟล์ที่อัปโหลด</div>
              </div>
              <div className="statValue">{files.length}</div>
            </div>

            <div className="stat">
              <div className="left">
                <div className={`dot ${rows.length>0 ? 'good' : (files.length>0 ? 'warn' : '')}`}></div>
                <div className="statLabel">แถว SKU ที่อ่านได้</div>
              </div>
              <div className="statValue">{stats.totalSkuRows}</div>
            </div>

            <div className="stat">
              <div className="left">
                <div className={`dot ${stats.uniqueRows>0 ? 'good' : ''}`}></div>
                <div className="statLabel">จำนวนแถว (รุ่น+สี)</div>
              </div>
              <div className="statValue">{stats.uniqueRows}</div>
            </div>

            <div className="stat">
              <div className="left">
                <div className={`dot ${stats.badSku===0 ? 'good' : 'warn'}`}></div>
                <div className="statLabel">SKU ที่ parse ไม่ได้</div>
              </div>
              <div className="statValue">{stats.badSku}</div>
            </div>

            {errors.length > 0 && (
              <button className="btn secondary" onClick={onDownloadErrors} disabled={loading}>
                ดาวน์โหลด Error Report (CSV)
              </button>
            )}

            <div className="kv">
              <div className="pill"><strong>Sizes:</strong> {SIZES.join(', ')}</div>
              <div className="pill"><strong>Missing:</strong> เติมเป็น 0</div>
              <div className="pill"><strong>Aggregation:</strong> Sum (กันกรณีซ้ำจากหลายไฟล์)</div>
            </div>
          </div>

          <div className="footerHelp">
            ถ้ามี SKU ที่ผิดรูปแบบ (เช่น ไม่มี - หรือไม่มีไซส์ท้ายสุด) ให้เปิดไฟล์ CSV error เพื่อตรวจสอบ แล้ว export ใหม่อีกครั้ง
          </div>
        </div>
      </div>
    </div>
  )
}
