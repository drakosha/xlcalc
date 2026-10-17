// node --expose-gc xlsx-child.mjs <rows> <xlcalc|exceljs|sheetjs>: the synthetic book of `rows` rows written by xlcalc's
// writeXlsx, read from its bytes by one library and written back from what it read, in a fresh process. The first read
// is a cold one whose retained heap is measured; one more warms up, five are timed, then the same for writing. The values
// read are checked against xlcalc's readXlsx outside the timer: dates as 1900 serials, a formula's cached result, rich
// text joined, a missing cell and an empty text alike, numbers within 1e-9 relative.
import { createRequire } from 'node:module'
import { pathToFileURL } from 'node:url'
import { resolve } from 'node:path'
import { CLOCK, equal, median } from '../hyperformula/shared.mjs'
import { dateToSerial, formatAddress, parseAddress } from '../../dist/model.js'
import { synthBook } from '../../dist/synth.js'
import { ERRORS } from './errors.mjs'

const [rows, lib] = process.argv.slice(2)
const root = process.env.XLCALC_BENCH_ROOT ? resolve(process.env.XLCALC_BENCH_ROOT) : resolve(import.meta.dirname, '../..')
const { compute } = await import(pathToFileURL(resolve(root, 'dist/index.js')))
const { readXlsx, writeXlsx } = await import(pathToFileURL(resolve(root, 'dist/xlsx.js')))
const need = createRequire(import.meta.url)
const ExcelJS = need('exceljs')
const XLSX = need('xlsx')
const send = (message) => process.send?.(message)
const phase = (name) => send({ type: 'phase', name })

const serial = (d) => d.getTime() / 86_400_000 + 25_569
/** A value as the comparison reads it. */
function plain(v) {
  if (v === undefined || v === null || v === '') return null
  if (v instanceof Date) return serial(v)
  if (typeof v === 'object' && 'date' in v) return dateToSerial(v.date)
  if (typeof v === 'object' && 'error' in v) return { error: ERRORS[v.error] ?? v.error }
  if (typeof v === 'object' && 'richText' in v) return plain(v.richText.map((r) => r.text).join(''))
  if (typeof v === 'object' && 'formula' in v) return plain(v.result)
  if (typeof v === 'object' && 'sharedFormula' in v) return plain(v.result)
  if (typeof v === 'object' && 'text' in v) return plain(v.text)
  return v
}

/** Each library's read, write and the values it read: `[sheet name, address, value]`. */
const LIBS = {
  xlcalc: {
    read: (data) => readXlsx(data),
    write: (read) => writeXlsx({ book: read.book, result: read.cached }),
    *cells(read) {
      for (const [i, sheet] of read.book.sheets.entries())
        for (const [addr, cell] of Object.entries(sheet.cells)) yield [sheet.name, addr, cell.formula !== undefined ? read.cached.sheets[i].cells[addr]?.value : cell.value]
    },
  },
  exceljs: {
    async read(data) {
      const wb = new ExcelJS.Workbook()
      await wb.xlsx.load(data)
      return wb
    },
    write: (wb) => wb.xlsx.writeBuffer(),
    *cells(wb) {
      const out = []
      for (const ws of wb.worksheets) ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => out.push([ws.name, cell.address, cell.value])))
      yield* out
    },
  },
  sheetjs: {
    read: (data) => XLSX.read(data),
    write: (wb) => XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }),
    *cells(wb) {
      for (const name of wb.SheetNames)
        for (const [addr, c] of Object.entries(wb.Sheets[name])) if (!addr.startsWith('!')) yield [name, addr, c.t === 'e' ? { error: c.w } : c.t === 'z' ? null : c.v]
    },
  },
}
const it = LIBS[lib]

/** The values xlcalc reads, by `sheet!address`: in a frame of their own, so that nothing of the read outlives it. */
async function valuesOf(data) {
  const out = new Map()
  for (const [sheet, addr, v] of LIBS.xlcalc.cells(await readXlsx(data))) out.set(`${sheet}!${addr}`, plain(v))
  return out
}

/** The time of an asynchronous call, as `measure` gives a synchronous one. */
async function timed(f) {
  const t = performance.now()
  const value = await f()
  return { value, wallMs: performance.now() - t }
}

try {
  phase('file')
  const { book } = synthBook({ seed: 1, rows: Number(rows) })
  const { data } = await writeXlsx({ book, result: compute(book, { clock: CLOCK }) })
  const expected = await valuesOf(data)
  phase('cold-read')
  global.gc()
  const heapBefore = process.memoryUsage().heapUsed
  let read = (await timed(() => it.read(data))).value
  global.gc()
  const retainedHeapMb = (process.memoryUsage().heapUsed - heapBefore) / 2 ** 20
  // the values read: every cell of xlcalc's read, and any other a library read with a value
  phase('verify')
  const got = new Map()
  for (const [sheet, addr, v] of it.cells(read)) got.set(`${sheet}!${formatAddress(parseAddress(addr).row, parseAddress(addr).col)}`, plain(v))
  const mismatches = []
  let differences = 0
  for (const k of new Set([...expected.keys(), ...got.keys()])) {
    const a = expected.get(k) ?? null, b = got.get(k) ?? null
    if (equal(a, b)) continue
    differences++
    if (mismatches.length < 5) mismatches.push({ cell: k, expected: a, actual: b })
  }
  const reads = [], writes = []
  for (let i = 0; i < 6; i++) {
    phase('read')
    read = undefined
    const t = await timed(() => it.read(data))
    read = t.value
    if (i) reads.push(t.wallMs)
  }
  let bytes = 0
  for (let i = 0; i < 6; i++) {
    phase('write')
    const t = await timed(() => it.write(read))
    bytes = (t.value.data ?? t.value).byteLength
    if (i) writes.push(t.wallMs)
  }
  send({ type: 'result', result: { status: differences ? 'mismatch' : 'ok', file: data.byteLength, cells: expected.size, readMs: median(reads), writeMs: median(writes), written: bytes, retainedHeapMb, reads, writes, differences, mismatches } })
} catch (error) {
  send({ type: 'result', result: { status: 'error', message: error.message } })
} finally {
  process.disconnect?.()
}
