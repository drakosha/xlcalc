/**
 * Structure operations, fills and pastes at size, from the review of 21.10 (speed): the work of a fill or a paste grows
 * with what it writes, not with the square of it or with the sheet times the tiles. Each test takes a size at which the
 * old code took seconds and the new takes a small part of the bound.
 */
import { describe, expect, it } from 'vitest'
import { createSession, editStructure, InputError, type BookInput, type StructureResult } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** Milliseconds `f` takes. */
function timed(f: () => void): number {
  const t = performance.now()
  f()
  return performance.now() - t
}

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it('a fill of a validated cell joins the pasted rules in linear time (speed-u123)', () => {
    const n = 20_000
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, validations: [{ range: 'A1', type: 'whole', operator: 'greaterThan', formula1: 'B1' }] }] }
    const s = createSession(book, { execution, clock, display: false })
    // the old join re-parsed the growing range per pasted rule: 16-18 s at 20,000 rows
    const ms = timed(() => expect(s.editStructure({ op: 'fillRange', sheet: 'S', range: 'A1', to: `A1:A${n}`, type: 'copy' }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    expect(s.getBook().sheets[0].validations).toEqual([{ range: `A1:A${n}`, type: 'whole', operator: 'greaterThan', formula1: 'B1' }])
  })

  it('a fill over rows with their own entries looks up only the lines of each tile (speed-u91)', () => {
    const n = 12_000
    const [cells, rows]: [BookInput['sheets'][0]['cells'], Record<string, { height: number }>] = [{ B1: { formula: 'A1*2' } }, {}]
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      rows[String(r)] = { height: 15 + (r % 3) }
    }
    const s = createSession({ sheets: [{ name: 'S', cells, rows }] }, { execution, clock, display: false })
    // the old fill walked every row entry per tile: 9-10 s at 12,000 rows
    const ms = timed(() => expect(s.editStructure({ op: 'fillRange', sheet: 'S', range: 'B1', to: `B1:B${n}`, type: 'copy' }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    expect(s.getCell('S', `B${n}`)).toMatchObject({ formula: `A${n}*2`, value: n * 2 })
  })

  it('deleting rows on a filtered sheet moves the cells once, not once per run of visible rows (speed-u90)', () => {
    const [n, runs] = [6000, 800]
    const [cells, rows]: [BookInput['sheets'][0]['cells'], Record<string, { hidden: true }>] = [{ C2: { formula: `SUM(A1:A${n})` } }, {}]
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      cells[`B${r}`] = { formula: `A${r}*2` }
    }
    for (let k = 1; k <= runs; k++) rows[String(2 * k)] = { hidden: true }
    const book: BookInput = { sheets: [{ name: 'S', cells, rows, filtered: true }, { name: 'T', cells: { A1: { formula: `S!A${n}+S!A3` } } }] }
    const s = createSession(book, { execution, clock, display: false })
    // the old delete rebuilt the book once per run: 8-10 s at 800 runs of 6,000 rows
    const ms = timed(() => expect(s.editStructure({ op: 'deleteRows', sheet: 'S', row: 1, count: 2 * runs }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    // the hidden rows stay, the visible ones above the end go
    expect(s.getCell('S', 'A1')).toMatchObject({ value: 2 })
    expect(s.getCell('S', 'B2')).toMatchObject({ formula: 'A2*2', value: 8 })
    expect(s.getCell('S', 'C1')).toMatchObject({ formula: `SUM(A1:A${n - runs})` })
    expect(s.getCell('T', 'A1')).toMatchObject({ formula: `'S'!A${n - runs}+'S'!#REF!`, value: { error: 'REF' } })
  })

  it('a paste over a filtered sheet walks its cells once, not once per run of visible rows (speed-u90)', () => {
    const [n, runs] = [34_000, 400]
    const [cells, rows]: [BookInput['sheets'][0]['cells'], Record<string, { hidden: true }>] = [{ E1: { formula: 'C1*3' } }, {}]
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      cells[`B${r}`] = { formula: `A${r}*2` }
    }
    for (let k = 1; k <= runs; k++) rows[String(2 * k)] = { hidden: true }
    const s = createSession({ sheets: [{ name: 'S', cells, rows, filtered: true }] }, { execution, clock, display: false })
    // the old paste rebuilt the sheet once per run: 4-5 s at 400 runs of 34,000 rows
    const ms = timed(() => expect(s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'E1' }, to: { sheet: 'S', range: `C1:C${2 * runs}` } }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    // the visible rows take the copy, the hidden ones keep what they had
    expect(s.getCell('S', 'C3')).toMatchObject({ formula: 'A3*3', value: 9 })
    expect(s.getCell('S', 'C2')?.formula).toBeUndefined()
    expect(s.getCell('S', `B${n}`)).toMatchObject({ formula: `A${n}*2`, value: 2 * n })
  })

  it('a paste on a sheet with a rule over many single cells joins them in linear time (speed-u124)', () => {
    const n = 300_000
    const range = Array.from({ length: n }, (_, r) => `A${r + 1}`).join(' ')
    const s = createSession({ sheets: [{ name: 'S', cells: { C1: { value: 1 } }, validations: [{ range, type: 'whole', operator: 'greaterThan', formula1: '0' }] }] }, { execution, clock, display: false })
    // the old join of touching rectangles spliced the list per join: 17 s at 300,000 cells
    const ms = timed(() => expect(s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'S', range: 'D1' } }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    expect(s.getBook().sheets[0].validations).toEqual([{ range: `A1:A${n}`, type: 'whole', operator: 'greaterThan', formula1: '0' }])
  })

  it('a merge across walks the cells of its range once, not once per row (speed-u93)', () => {
    const n = 20_000
    const cells: BookInput['sheets'][0]['cells'] = {}
    for (let r = 1; r <= n; r++) for (const c of 'ABDEF') cells[`${c}${r}`] = { value: r }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, { execution, clock, display: false })
    // the old merge walked the sheet's cells for each row: 5-6 s at 20,000 rows of 100,000 cells
    const ms = timed(() => expect(s.editStructure({ op: 'mergeCells', sheet: 'S', range: `A1:C${n}`, across: true }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    const sheet = s.getBook().sheets[0]
    expect(sheet.merges?.length).toBe(n)
    expect([sheet.cells.A7?.value, sheet.cells.B7, sheet.cells.D7?.value]).toEqual([7, undefined, 7])
  })

  describe('editStructure checks the cells it was given, not those it made (speed-u95)', () => {
    const book = (): BookInput => ({ styles: [{}, { font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1, style: 1 }, A2: { formula: '=A1+1' }, B1: { style: 1 } } }, { name: 'T', cells: { A1: { formula: 'S!A2*2' } } }] })
    const done = (r: StructureResult) => {
      if (!r.ok) throw new Error(r.message)
      return r.book
    }

    it('the book it gave, given again, comes out as a copy of it does', () => {
      const once = done(editStructure(book(), { op: 'insertRows', sheet: 'S', row: 1 }))
      const op = { op: 'copyRange', from: { sheet: 'S', range: 'A2:B3' }, to: { sheet: 'S', range: 'C5' } } as const
      expect(done(editStructure(once, op))).toEqual(done(editStructure(structuredClone(once), op)))
    })

    it("the caller's own cells are checked on every call, changed in place too", () => {
      const b = book()
      done(editStructure(b, { op: 'insertRows', sheet: 'S', row: 1 }))
      b.sheets[0].cells.A9 = { value: Number.NaN }
      expect(() => editStructure(b, { op: 'insertRows', sheet: 'S', row: 1 })).toThrow(InputError)
    })

    it('cells it made are checked again against fewer styles', () => {
      const once = done(editStructure(book(), { op: 'insertRows', sheet: 'S', row: 1 }))
      expect(() => editStructure({ ...once, styles: [{}] }, { op: 'insertRows', sheet: 'S', row: 1 })).toThrow(/style/)
    })
  })
})
