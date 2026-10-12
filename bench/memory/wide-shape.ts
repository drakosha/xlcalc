/**
 * The memory and time of a book of the wide shape, made up, no data of it: sheet G of 344 rows from A to XFD of cells that
 * hold only a look, two looks, the second in the same 2,728 columns of every row (columns R, T, AD, AF and every 12th
 * after them), rows in threes with two empty rows between, a little content above; a small second sheet reading it.
 * It runs over the built package (`npm run build`):
 *
 *   npx tsx bench/memory/wide-shape.ts write <out.xlsx>    the file, as writeXlsx writes it
 *   node --expose-gc --import tsx bench/memory/wide-shape.ts measure <file.xlsx>
 *
 * `measure` prints the heap a fresh process holds more after a full collection: after `readXlsx`, then with a session
 * over its book, the time each took, and the memory of the process before and after a pool's thread opens the book.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import type { BookInput, CellInput } from '../../src/contract.js'

const dist = (file: string) => new URL(`../../dist/${file}`, import.meta.url).href
const COLS = 16_384

/** The book: entries of a look alone over sheet G, as a book held them before the looks; a reader puts them in the looks. */
function wideShape(): BookInput {
  const colName = (c: number) => {
    let s = ''
    for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
    return s
  }
  const second = new Set<number>()
  for (let c = 18; c <= COLS; c += 12) {
    second.add(c)
    if (c + 2 <= COLS) second.add(c + 2)
  }
  const rows: number[] = []
  for (let k = 0; rows.length < 344; k++) for (let j = 0; j < 3 && rows.length < 344; j++) rows.push(20 + k * 5 + j)
  const names = Array.from({ length: COLS + 1 }, (_, c) => (c ? colName(c) : ''))
  const [a, b] = [{ style: 0, numberFormat: '#,##0' }, { style: 1, numberFormat: '0.0%' }]
  const cells: Record<string, CellInput> = {}
  for (const r of rows) for (let c = 1; c <= COLS; c++) cells[names[c] + r] = second.has(c) ? b : a
  for (let r = 1; r <= 15; r++) for (let c = 1; c <= 20; c++) cells[names[c] + r] = r === 1 ? { value: 'h' + c, style: 0 } : c === 20 ? { formula: `SUM(A${r}:S${r})` } : { value: r * c }
  return { styles: [{ font: { bold: true } }, { fill: { color: '#FFF2CC' } }], sheets: [{ name: 'G', cells }, { name: 'Data', cells: { A1: { value: 1 }, A2: { formula: 'G!T2*2' } } }] }
}

/** The heap in MB after full collections (`--expose-gc`). */
function heap(): number {
  const gc = (globalThis as { gc?: () => void }).gc
  if (!gc) throw new Error('run with node --expose-gc')
  gc()
  gc()
  return process.memoryUsage().heapUsed / 2 ** 20
}

async function measure(file: string) {
  const { readXlsx } = (await import(dist('xlsx.js'))) as typeof import('../../src/xlsx.js')
  const { createSession } = (await import(dist('index.js'))) as typeof import('../../src/index.js')
  const { createCalcPool } = (await import(dist('worker/index.js'))) as typeof import('../../src/worker/index.js')
  const bytes = new Uint8Array(readFileSync(file))
  const round = (x: number) => Math.round(x)
  const h0 = heap()
  let t = performance.now()
  const read = await readXlsx(bytes)
  const readMs = performance.now() - t
  const h1 = heap()
  t = performance.now()
  const s = createSession(read.book, { clock: { now: '2026-10-12T12:00:00Z', seed: 1 } })
  const sessionMs = performance.now() - t
  const h2 = heap()
  const looks = read.book.sheets.reduce((n, x) => n + (x.looks?.length ?? 0), 0)
  const entries = read.book.sheets.reduce((n, x) => n + Object.keys(x.cells).length, 0)
  s.close()
  // a pool's thread: the process's memory (resident, and the heaps of all threads) before and after it opens the book
  const pool = createCalcPool({ threads: 1 })
  const mem = () => round(process.memoryUsage().rss / 2 ** 20)
  await (await pool.createSession({ sheets: [{ name: 'S', cells: {} }] })).close()
  const before = mem()
  t = performance.now()
  const ps = await pool.createSession(read.book)
  const poolMs = performance.now() - t
  const after = mem()
  await ps.close()
  await pool.close()
  console.log(JSON.stringify({ entries, looks, readMs: round(readMs), readMb: round(h1 - h0), sessionMs: round(sessionMs), sessionMb: round(h2 - h1), poolMs: round(poolMs), poolRssMb: [before, after] }))
}

const [what, file] = process.argv.slice(2)
if (what === 'write') {
  const { writeXlsx } = (await import(dist('xlsx.js'))) as typeof import('../../src/xlsx.js')
  const { data } = await writeXlsx({ book: wideShape() }, { values: 'omit' })
  writeFileSync(file, data)
  console.log(`${file}: ${data.length} bytes`)
} else if (what === 'measure') await measure(file)
else throw new Error('expected write <out.xlsx> or measure <file.xlsx>')
