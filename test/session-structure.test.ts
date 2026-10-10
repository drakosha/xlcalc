/**
 * Structure operations, styles and the layout in a session: `restructure` gives the book and the values a full
 * computation of the restructured book gives, and its changes turn the old values moved by `mapCell` into the
 * new ones; style and layout changes recompute nothing; `book()` is the current input.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, openSession, restructure, type BookInput, type CellValue, type ComputeResult, type StructureOp } from '../src/index.js'
import { indexToCol } from '../src/model.js'
import { mulberry32, S, synthBook } from '../src/synth.js'
import { randomBook, SHEETS } from './seeded-books.js'

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

describe('session.restructure', () => {
  it('seeded books: the same book and hash as restructure and compute; changes and mapCell give the new values', () => {
    for (let seed = 1; seed <= 12; seed++) {
      const book = randomBook(seed)
      const session = openSession(book, { clock })
      let expected = book
      for (const op of ops(seed * 101)) {
        const before = session.result()
        const pure = restructure(expected, op)
        const r = session.restructure(op)
        if (!pure.ok || !r.ok) {
          expect(r, `seed ${seed} ${op.op}`).toEqual(pure)
          continue
        }
        expected = pure.book
        expect(r.book, `seed ${seed} ${op.op}`).toEqual(pure.book)
        expect(r.book).toBe(session.book())
        const full = compute(pure.book, { clock })
        expect(session.result().hash, `seed ${seed} ${op.op}`).toBe(full.hash)
        expect(shown(before, r), `seed ${seed} ${op.op}`).toEqual(valuesOf(full))
      }
      session.close()
    }
  })

  it('an insert into the data sheet of the synthetic book changes no value; rebuilding is within the cold budget', () => {
    const { book } = synthBook({ seed: 1, rows: process.env.FORMULA_BENCH === '1' ? 3400 : 340 })
    const session = openSession(book, { clock })
    const t0 = performance.now()
    const r = session.restructure({ op: 'insertRows', sheet: S.data, row: 170, count: 1 })
    const ms = performance.now() - t0
    if (process.env.FORMULA_BENCH === '1') console.log(`session.restructure, insert a row into ${S.data}: ${ms.toFixed(0)} ms`)
    expect(r.ok && r.changes).toEqual([])
    expect(ms).toBeLessThan(1500)
    expect(session.result().hash).toBe(compute(session.book(), { clock }).hash)
  }, 60_000)

  it('a refused operation changes nothing; a live session takes a new clock', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'RAND()' }, A2: { value: 1 } } }] }
    const s = openSession(book, { clock })
    expect(s.restructure({ op: 'deleteSheet', sheet: 'S' })).toMatchObject({ ok: false, code: 'last_sheet' })
    expect(s.book()).toEqual(book)
    const m = s.restructure({ op: 'insertRows', sheet: 's', row: 1 })
    expect(m.ok && m.mapCell({ sheet: 's', cell: 'A2' })).toEqual({ sheet: 'S', cell: 'A3' })
    expect(() => m.ok && m.mapCell({ sheet: 'S', cell: 'a2' })).toThrow(InputError)
    const live = openSession(book, { live: true })
    const c = live.clock
    const r = live.restructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(live.clock.seed).not.toBe(c.seed)
    expect(r.ok && r.changes.map((x) => x.cell)).toEqual(['A2'])
  })
})

describe('styles and layout in a session', () => {
  const book = (): BookInput => ({
    sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*2', style: 0 } }, columns: { 'A:Z': { width: 15 } }, merges: ['C1:D1'] }],
    styles: [{ font: { bold: true } }],
  })

  it('a cell edit that keeps the value and formula recomputes nothing and returns no changes', () => {
    const s = openSession(book(), { live: true, display: false })
    const c = s.clock
    expect(s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: 'A1*2', format: '0.0' } }, { sheet: 'S', cell: 'B9', input: { style: 0 } }])).toEqual([])
    expect(s.clock).toEqual(c)
    expect(s.book().sheets[0].cells).toEqual({ A1: { value: 2 }, A2: { formula: 'A1*2', format: '0.0' }, B9: { style: 0 } })
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3, style: 0 } }, { sheet: 'S', cell: 'B9', input: null }]).map((x) => x.cell)).toEqual(['A1', 'A2'])
    expect(s.book().sheets[0].cells.B9).toBeUndefined()
  })

  it('setStyles and setLayout change the book, not the values; spans are split where an edit cuts them', () => {
    const s = openSession(book(), { clock })
    const before = s.result()
    s.setStyles([{ font: { bold: true } }, { fill: { color: '#FFC000' } }])
    s.setLayout([
      { sheet: 'S', columns: 'C:D', set: { width: 30, style: 1 } },
      { sheet: 's', rows: '2', set: { height: 24 } },
      { sheet: 'S', columns: 'D', set: { width: null } },
      { sheet: 'S', merges: ['C1:D2'] },
    ])
    const sheet = s.book().sheets[0]
    expect(sheet.columns).toEqual({ 'A:B': { width: 15 }, C: { width: 30, style: 1 }, D: { style: 1 }, 'E:Z': { width: 15 } })
    expect(sheet.rows).toEqual({ 2: { height: 24 } })
    expect(sheet.merges).toEqual(['C1:D2'])
    const after = s.result()
    expect(after.hash).toBe(before.hash)
    expect(after.inputHash).not.toBe(before.inputHash)
    expect(after.inputHash).toBe(compute(s.book(), { clock }).inputHash)
    s.setLayout([{ sheet: 'S', merges: [] }, { sheet: 'S', rows: '2', set: { height: null } }])
    expect(s.book().sheets[0].merges).toBeUndefined()
    expect(s.book().sheets[0].rows).toBeUndefined()
  })

  it('wrong styles and layout edits throw InputError and change nothing', () => {
    const s = openSession(book(), { clock })
    expect(() => s.setStyles([])).toThrow(new InputError('styles', 'S!A2 uses style 0, the new table has 0'))
    expect(() => s.setLayout([{ sheet: 'S', columns: 'B', set: { width: 10 } }, { sheet: 'S', columns: 'C', set: { style: 1 } }])).toThrow('edits[1].set.style: expected the index of a style of the book, 0 to 0, got number 1')
    expect(() => s.setLayout([{ sheet: 'T', rows: '1', set: {} }])).toThrow('edits[0].sheet: unknown sheet "T"')
    expect(() => s.setLayout([{ sheet: 'S', merges: ['A1:B2', 'B2:C3'] }])).toThrow('edits[0].merges[1]: B2:C3 overlaps A1:B2')
    expect(() => s.setLayout([{ sheet: 'S', cols: 'A', set: {} } as never])).toThrow('edits[0]: unknown field "cols" (expected sheet, columns, set) (one of columns, rows, merges, filtered or frozen)')
    expect(() => s.setCells([{ sheet: 'S', cell: 'A1', input: { style: 1 } }])).toThrow('edits[0].input.style: expected the index of a style of the book, 0 to 0, got number 1')
    expect(s.book()).toEqual(book())
  })
})
