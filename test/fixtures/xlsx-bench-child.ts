/**
 * Speed of the xlsx subpath in a fresh process (for `xlsx-budget.test.ts`; `--expose-gc`): a book of `rows` × 40
 * values (text, numbers, booleans, decimals in turn) written and read 3 times each, the median of each and the heap
 * a read book holds; then the synthetic book (3,400 rows) with its result written and read the same way.
 */
const rows = Number(process.argv[2] ?? 20_000)
const { libUrl } = await import('./target.js')
const { compute } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')
const { synthBook } = (await import(libUrl('synth.js'))) as typeof import('../../src/synth.js')
const { readXlsx, writeXlsx } = (await import(libUrl('xlsx.js'))) as typeof import('../../src/xlsx.js')
const { formatAddress } = (await import(libUrl('model.js'))) as typeof import('../../src/model.js')
import type { BookInput, CellInput } from '../../src/contract.js'

const gc = () => (globalThis as { gc?: () => void }).gc?.()
const median = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]

async function measure(book: BookInput, opts: Parameters<typeof writeXlsx>[1]) {
  const write: number[] = []
  const read: number[] = []
  let data: Uint8Array = new Uint8Array()
  for (let i = 0; i < 3; i++) {
    const t = performance.now()
    data = (await writeXlsx(book, opts)).data
    write.push(performance.now() - t)
  }
  gc()
  const base = process.memoryUsage().heapUsed
  let last: Awaited<ReturnType<typeof readXlsx>> | undefined
  for (let i = 0; i < 3; i++) {
    last = undefined
    gc()
    const t = performance.now()
    last = await readXlsx(data)
    read.push(performance.now() - t)
  }
  gc()
  // the heap the last read book holds
  const heapMB = Math.round((process.memoryUsage().heapUsed - base) / 1e6)
  return { writeMs: Math.round(median(write)), readMs: Math.round(median(read)), sizeMB: Math.round(data.length / 1e4) / 100, heapMB, sheets: last!.book.sheets.length }
}

const cells: Record<string, CellInput> = {}
for (let r = 0; r < rows; r++)
  for (let c = 0; c < 40; c++) cells[formatAddress(r, c)] = { value: c % 4 === 0 ? `text ${r % 977} ${c}` : c % 4 === 1 ? r * 1.5 + c : c % 4 === 2 ? r % 2 === 0 : r + c / 1000 }
const values = await measure({ sheets: [{ name: 'Values', cells }] }, {})
const synth = synthBook({ seed: 1 }).book
const synthResult = await measure(synth, { result: compute(synth, { clock: { now: '2026-10-07T12:00:00Z', seed: 7 } }) })
console.log(JSON.stringify({ rows, values, synth: synthResult }))
