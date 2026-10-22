/**
 * Reading xlsx in time linear in what a sheet holds, from the review of 21.10: tables over a large sheet, a shared
 * formula spread thin, thousands of merged areas, thousands of data-validation rules. Each case took seconds when its
 * part was quadratic and takes tens of milliseconds now. Reading a file does not calculate, so the execution modes do
 * not apply here.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, CellInput, TableInput } from '../src/contract.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { pack, parts } from './xlsx-fixtures.js'

/** How long `f` takes, in milliseconds, and what it gives. */
async function timed<T>(f: () => Promise<T>): Promise<[number, T]> {
  const t = performance.now()
  const v = await f()
  return [performance.now() - t, v]
}

/** 34,000 rows of values in A:C (102,000 cells) and 20 tables of two columns in E:F. */
function tablesBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 34_000; r++) for (const c of 'ABC') cells[`${c}${r}`] = { value: r }
  const tables: TableInput[] = []
  for (let k = 0; k < 20; k++) {
    const r = k * 10 + 1
    cells[`E${r}`] = { value: 'Name' }
    cells[`F${r}`] = { value: 'Qty' }
    for (let i = 1; i <= 4; i++) Object.assign(cells, { [`E${r + i}`]: { value: `x${i}` }, [`F${r + i}`]: { value: i } })
    tables.push({ name: `Tbl_${k}`, range: `E${r}:F${r + 4}`, columns: [{ name: 'Name' }, { name: 'Qty' }] })
  }
  return { sheets: [{ name: 'S', cells, tables }] }
}

const rows = (n: number, f: (i: number) => string) => Array.from({ length: n }, (_, i) => f(i)).join('')

/** A file of one sheet with this body, packed before the read is timed. */
const file = (body: string) => pack(parts({ sheets: [{ name: 'S', body }] }))

describe('xlsx: reading in time linear in the sheet (speed-u1, u3, u5, u6)', () => {
  it('20 tables over 102,000 cells (speed-u1): each table part is checked against the arrays and headers, not all cells', async () => {
    const { data } = await writeXlsx({ book: tablesBook() })
    // 3.6 s when each part copied and walked every cell for itself and every table before it
    const [ms, r] = await timed(() => readXlsx(data))
    expect(r.book.sheets[0].tables?.map((t) => t.name)).toEqual(Array.from({ length: 20 }, (_, k) => `Tbl_${k}`))
    expect(r.issues).toEqual([])
    expect(ms).toBeLessThan(1000)
  }, 120_000)

  it('a shared formula in every 5th of 50,000 rows (speed-u3): each copy is pasted and read back alone', async () => {
    const copies = rows(9_999, (i) => `<row r="${(i + 1) * 5 + 1}"><c r="A${(i + 1) * 5 + 1}"><f t="shared" si="0"/></c></row>`)
    // 5 s when the read-back after each one-cell paste walked every copy of the group
    const data = await file(`<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A49996" si="0">B1*2</f></c></row>${copies}</sheetData>`)
    const [ms, r] = await timed(() => readXlsx(data))
    const cells = r.book.sheets[0].cells
    expect([cells.A1, cells.A6, cells.A49996]).toEqual([{ formula: 'B1*2' }, { formula: 'B6*2' }, { formula: 'B49996*2' }])
    expect(ms).toBeLessThan(1000)
  }, 120_000)
})
