/**
 * Structure operations, styles and the layout in a session: `editStructure` gives the book and the values a full
 * computation of the restructured book gives, and its changes turn the old values moved by `mapCell` into the
 * new ones; style and layout changes recompute nothing; `getBook()` is the current input.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, createSession, editStructure, type BookInput, type CellValue, type ComputeResult, type StructureOp } from '../src/index.js'
import { indexToCol } from '../src/model.js'
import { mulberry32, S, synthBook } from '../src/synth.js'
import { cellsOf } from './looks-book.js'
import { randomBook, SHEETS } from './seeded-books.js'
import { withHeld } from './session-edits.js'

const clock = { now: '2026-10-06T12:00:00Z', seed: 7 }

const valuesOf = (r: ComputeResult) => new Map(r.sheets.flatMap((s) => Object.entries(s.cells).filter(([, c]) => c.value !== null).map(([a, c]): [string, CellValue] => [`${s.name}!${a}`, c.value])))

/** Old values moved by mapCell, then the changes applied: what an interface shows after the operation. */
function shown(before: ComputeResult, r: { mapCell: (ref: { sheet: string; cell: string }) => { sheet: string; cell: string } | null; changes: { sheet: string; cell: string; value: CellValue }[] }) {
  const out = new Map<string, CellValue>()
  for (const [k, v] of valuesOf(before)) {
    const [sheet, cell] = k.split('!')
    const to = r.mapCell({ sheet, cell })
    if (to) out.set(`${to.sheet}!${to.cell}`, v)
  }
  for (const c of r.changes) {
    if (c.value === null) out.delete(`${c.sheet}!${c.cell}`)
    else out.set(`${c.sheet}!${c.cell}`, c.value)
  }
  return out
}

function ops(seed: number): StructureOp[] {
  const rnd = mulberry32(seed)
  const int = (n: number) => Math.floor(rnd() * n)
  const sheet = () => SHEETS[int(3)]
  const range = () => `${indexToCol(int(14))}${1 + int(20)}:${indexToCol(14 + int(3))}${21 + int(5)}`
  return [
    { op: 'insertRows', sheet: sheet(), row: 1 + int(22), count: 1 + int(3) },
    { op: 'deleteRows', sheet: sheet(), row: 1 + int(22), count: 1 + int(3) },
    { op: 'insertColumns', sheet: sheet(), column: indexToCol(int(15)), count: 1 + int(2) },
    { op: 'deleteColumns', sheet: sheet(), column: indexToCol(int(15)), count: 1 + int(2) },
    { op: 'moveRange', from: { sheet: sheet(), range: range() }, to: { sheet: sheet(), cell: `${indexToCol(int(16))}${1 + int(30)}` } },
    { op: 'copyRange', from: { sheet: sheet(), range: range() }, to: { sheet: sheet(), range: `${indexToCol(int(16))}${1 + int(30)}` } },
    { op: 'renameSheet', sheet: sheet(), name: `R${seed}` },
  ]
}

describe('session.editStructure', () => {
  it('seeded books: the same book and hash as editStructure and compute; changes and mapCell give the new values', () => {
    for (let seed = 1; seed <= 12; seed++) {
      const book = randomBook(seed)
      const session = createSession(book, { clock })
      let expected = book
      for (const op of ops(seed * 101)) {
        const before = session.getResult()
        const pure = editStructure(expected, op)
        const r = session.editStructure(op)
        if (!pure.ok || !r.ok) {
          expect(r, `seed ${seed} ${op.op}`).toEqual(pure)
          continue
        }
        // a cell an operation brings into a cycle's way holds what it showed (`CellInput.held`)
        expected = withHeld(pure.book, session.getBook())
        expect(session.getBook(), `seed ${seed} ${op.op}`).toEqual(expected)
        expect(r).not.toHaveProperty('book')
        const full = compute(expected, { clock })
        expect(session.getResult().hash, `seed ${seed} ${op.op}`).toBe(full.hash)
        expect(shown(before, r), `seed ${seed} ${op.op}`).toEqual(valuesOf(full))
      }
      session.close()
    }
  })

  it('an insert into the data sheet of the synthetic book changes no value, only the text of the formulas over it; rebuilding is within the cold budget', () => {
    const { book } = synthBook({ seed: 1, rows: process.env.FORMULA_BENCH === '1' ? 3400 : 340 })
    const session = createSession(book, { clock })
    const t0 = performance.now()
    const r = session.editStructure({ op: 'insertRows', sheet: S.data, row: 170, count: 1 })
    const ms = performance.now() - t0
    if (process.env.FORMULA_BENCH === '1') console.log(`session.editStructure, insert a row into ${S.data}: ${ms.toFixed(0)} ms`)
    // the formulas over the data sheet read one row more: listed with their text, none with another value
    expect(r.ok && r.changes.filter((c) => c.formula === undefined)).toEqual([])
    expect(r.ok && r.changes.length).toBeGreaterThan(0)
    expect(ms).toBeLessThan(1500)
    expect(session.getResult().hash).toBe(compute(session.getBook(), { clock }).hash)
  }, 60_000)

  it('a fill pastes its source once per offset, not the sheet: one row filled down 2,000 rows of the synthetic book is within the cold budget', () => {
    // the looks of a paste once walked every cell of the source sheet for each of the 1,998 offsets (11 s here)
    const { book } = synthBook({ seed: 1, rows: 3400 })
    const session = createSession(book, { clock })
    const t0 = performance.now()
    const r = session.editStructure({ op: 'copyRange', from: { sheet: S.data, range: 'A2:T2' }, to: { sheet: S.data, range: 'A3:T2000' } })
    const ms = performance.now() - t0
    expect(r.ok).toBe(true)
    expect(ms).toBeLessThan(1500)
    expect(session.getResult().hash).toBe(compute(session.getBook(), { clock }).hash)
  }, 60_000)

  it('a refused operation changes nothing; a live session takes a new clock', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'RAND()' }, A2: { value: 1 } } }] }
    const s = createSession(book, { clock })
    expect(s.editStructure({ op: 'deleteSheet', sheet: 'S' })).toMatchObject({ ok: false, code: 'last_sheet' })
    expect(s.getBook()).toEqual(book)
    const m = s.editStructure({ op: 'insertRows', sheet: 's', row: 1 })
    expect(m.ok && m.mapCell({ sheet: 's', cell: 'A2' })).toEqual({ sheet: 'S', cell: 'A3' })
    expect(() => m.ok && m.mapCell({ sheet: 'S', cell: 'a2' })).toThrow(InputError)
    const live = createSession(book, { live: true })
    const c = live.clock
    const r = live.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(live.clock.seed).not.toBe(c.seed)
    expect(r.ok && r.changes.map((x) => x.cell)).toEqual(['A2'])
  })
})

describe('styles and layout in a session', () => {
  const book = (): BookInput => ({
    sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*2', style: 0 } }, columns: { 'A:Z': { width: 15 } }, merges: ['C1:D1'] }],
    styles: [{ font: { bold: true } }],
  })

  it('a cell edit that keeps the value and formula recomputes nothing and returns the cells with their new look', () => {
    const s = createSession(book(), { live: true, display: false })
    const c = s.clock
    expect(s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: 'A1*2', numberFormat: '0.0', style: null } }, { sheet: 'S', cell: 'B9', input: { style: 0 } }]).changes).toEqual([
      { sheet: 'S', cell: 'A2', value: 4, style: null, numberFormat: '0.0', numberFormatCode: '0.0' },
      { sheet: 'S', cell: 'B9', value: null, style: 0 },
    ])
    expect(s.clock).toEqual(c)
    expect(cellsOf(s.getBook().sheets[0])).toEqual({ A1: { value: 2 }, A2: { formula: 'A1*2', numberFormat: '0.0' }, B9: { style: 0 } })
    // a cell cleared keeps its look; replaced by nothing, it is gone
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3, style: 0 } }, { sheet: 'S', cell: 'B9', input: null }]).changes.map((x) => x.cell)).toEqual(['A1', 'A2'])
    expect(cellsOf(s.getBook().sheets[0]).B9).toEqual({ style: 0 })
    s.replaceCells([{ sheet: 'S', cell: 'B9', input: null }])
    expect(cellsOf(s.getBook().sheets[0]).B9).toBeUndefined()
  })

  it('setStyles and setLayout change the book, not the values; spans are split where an edit cuts them', () => {
    const s = createSession(book(), { clock })
    const before = s.getResult()
    s.setStyles([{ font: { bold: true } }, { fill: { color: '#FFC000' } }])
    s.setLayout([
      { sheet: 'S', columns: 'C:D', set: { width: 30, style: 1 } },
      { sheet: 's', rows: '2', set: { height: 24 } },
      { sheet: 'S', columns: 'D', set: { width: null } },
      { sheet: 'S', merges: ['C1:D2'] },
    ])
    const sheet = s.getBook().sheets[0]
    expect(sheet.columns).toEqual({ 'A:B': { width: 15 }, C: { width: 30, style: 1 }, D: { style: 1 }, 'E:Z': { width: 15 } })
    expect(sheet.rows).toEqual({ 2: { height: 24 } })
    expect(sheet.merges).toEqual(['C1:D2'])
    const after = s.getResult()
    expect(after.hash).toBe(before.hash)
    expect(after.inputHash).not.toBe(before.inputHash)
    expect(after.inputHash).toBe(compute(s.getBook(), { clock }).inputHash)
    s.setLayout([{ sheet: 'S', merges: [] }, { sheet: 'S', rows: '2', set: { height: null } }])
    expect(s.getBook().sheets[0].merges).toBeUndefined()
    expect(s.getBook().sheets[0].rows).toBeUndefined()
  })

  it('wrong styles and layout edits throw InputError and change nothing', () => {
    const s = createSession(book(), { clock })
    expect(() => s.setStyles([])).toThrow(new InputError('styles', 'S!A2 uses style 0, the new table has 0'))
    expect(() => s.setLayout([{ sheet: 'S', columns: 'B', set: { width: 10 } }, { sheet: 'S', columns: 'C', set: { style: 1 } }])).toThrow('edits[1].set.style: expected the index of a style of the workbook, 0 to 0, got number 1')
    expect(() => s.setLayout([{ sheet: 'T', rows: '1', set: {} }])).toThrow('edits[0].sheet: unknown sheet "T"')
    expect(() => s.setLayout([{ sheet: 'S', merges: ['A1:B2', 'B2:C3'] }])).toThrow('edits[0].merges[1]: B2:C3 overlaps A1:B2')
    expect(() => s.setLayout([{ sheet: 'S', cols: 'A', set: {} } as never])).toThrow('edits[0]: unknown field "cols" (expected sheet, columns, set) (one of columns, rows, merges, filtered, frozen, conditionalFormats, validations or outline)')
    expect(() => s.setCells([{ sheet: 'S', cell: 'A1', input: { style: 1 } }])).toThrow('edits[0].input.style: expected the index of a style of the workbook, 0 to 0, got number 1')
    expect(s.getBook()).toEqual(book())
  })
})
