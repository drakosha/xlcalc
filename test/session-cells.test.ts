/**
 * `Session.cells(sheet, range)`: the cells of a rectangle are the slice of `result()` over it, with each cell's own
 * style, in a synchronous session, in the same thread and in a thread, with closures and eval: the case books, the
 * tables of the table sets, a book with styles, rectangles small (read cell by cell) and large (whole columns, whole
 * rows, the whole sheet), before and after edits.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { buildTableCases } from '../bench/coverage/tables/ops.js'
import { createInlineCalc, InputError, openSession, type BookInput, type CellView, type ComputeResult, type Execution } from '../src/index.js'
import { parseAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'
import { synthBook } from '../src/synth.js'
import { BOOKS, CLOCK } from './async-books.js'
import { pool } from './fixtures/pool.js'
import { edits } from './session-edits.js'

/** The slice of a result over a rectangle, with the cells' own styles from the book. */
function slice(r: ComputeResult, book: BookInput, sheet: string, range: string): Record<string, CellView> {
  const x = parseRange(range)!
  const input = book.sheets.find((s) => s.name === sheet)!.cells
  const out: Record<string, CellView> = {}
  for (const [a, c] of Object.entries(r.sheets.find((s) => s.name === sheet)!.cells)) {
    const { row, col } = parseAddress(a)
    if (row < x.r1 || row > x.r2 || col < x.c1 || col > x.c2) continue
    const style = input[a]?.style
    out[a] = style === undefined ? c : { ...c, style }
  }
  return out
}

/** Rectangles of a sheet: small ones over its corner and its middle, a column and a row, whole columns, whole rows, everything. */
const RANGES = ['A1', 'A1:D12', 'B3:H40', 'C1:C1048576', 'A5:XFD5', 'B:D', '2:30', 'A:XFD']

const styled: BookInput = {
  styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }],
  sheets: [{ name: 'S', cells: { A1: { value: 1, style: 0 }, A2: { formula: 'A1*2', style: 1, format: '0.00' }, B1: { formula: 'SEQUENCE(3,2)' }, C5: { style: 1 }, D9: { value: 'x', meta: { id: 7 } } } }],
}
const BOOKS_HERE: [string, BookInput][] = [...BOOKS, ['a book with styles and a spill', styled], ...buildTableCases().slice(0, 6).map((c): [string, BookInput] => [`the table case ${c.id}`, c.book])]

type Cells = (sheet: string, range: string) => Promise<Record<string, CellView>>

/** Every rectangle of every sheet against the slice of the same session's result. */
async function agree(name: string, book: BookInput, cells: Cells, result: () => Promise<ComputeResult>, current: () => Promise<BookInput>) {
  const r = await result()
  const b = await current()
  for (const sheet of b.sheets) for (const range of RANGES) expect(await cells(sheet.name, range), `${name}: ${sheet.name}!${range}`).toEqual(slice(r, b, sheet.name, range))
}

const pools: { close(): Promise<void> }[] = []
afterAll(() => Promise.all(pools.map((p) => p.close())))

describe.each(['closures', 'eval'] as const)('cells of a rectangle, execution %s', (execution: Execution) => {
  it('a synchronous session', async () => {
    for (const [name, book] of BOOKS_HERE) {
      const s = openSession(book, { clock: CLOCK, execution })
      await agree(name, book, async (sh, rg) => s.cells(sh, rg), async () => s.result(), async () => s.book())
    }
  }, 120_000)

  it('in the same thread', async () => {
    const calc = createInlineCalc({ sliceMs: 2 })
    for (const [name, book] of BOOKS_HERE) {
      const s = await calc.openSession(book, { clock: CLOCK, execution })
      await agree(name, book, s.cells, s.result, s.book)
    }
  }, 120_000)

  it('in a thread', async () => {
    const p = await pool({ threads: 2 })
    pools.push(p)
    for (const [name, book] of BOOKS_HERE) {
      const s = await p.openSession(book, { clock: CLOCK, execution })
      await agree(name, book, s.cells, s.result, s.book)
    }
  }, 120_000)
})

describe('cells of a rectangle', () => {
  it('after edits, and without the display', async () => {
    const synth = synthBook({ seed: 1, rows: 340 })
    const s = openSession(synth.book, { clock: CLOCK })
    const quiet = openSession(synth.book, { clock: CLOCK, display: false })
    const p = await pool({ threads: 1 })
    pools.push(p)
    const t = await p.openSession(synth.book, { clock: CLOCK })
    for (const batch of edits(2, 12, synth)) {
      s.setCells(batch)
      quiet.setCells(batch)
      await t.setCells(batch)
    }
    await agree('the synthetic book after edits', synth.book, async (sh, rg) => s.cells(sh, rg), async () => s.result(), async () => s.book())
    await agree('the synthetic book in a thread after edits', synth.book, t.cells, t.result, t.book)
    await agree('without the display', synth.book, async (sh, rg) => quiet.cells(sh, rg), async () => quiet.result(), async () => quiet.book())
  }, 120_000)

  it('the sheet and the range are checked as input', () => {
    const s = openSession(styled, { clock: CLOCK })
    expect(() => s.cells('Nope', 'A1')).toThrow(InputError)
    expect(() => s.cells('S', 'A0')).toThrow('range: expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got "A0"')
    expect(s.cells('s', 'B2:B3')).toEqual({ B2: { value: 3, text: '3', spillOf: 'B1' }, B3: { value: 5, text: '5', spillOf: 'B1' } })
  })
})
