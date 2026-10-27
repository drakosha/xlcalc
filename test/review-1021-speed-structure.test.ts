/**
 * Structure operations, fills and pastes at size, from the review of 21.10 (speed): the work of a fill or a paste grows
 * with what it writes, not with the square of it or with the sheet times the tiles. Each test takes a size at which the
 * old code took seconds and the new takes a small part of the bound.
 */
import { describe, expect, it } from 'vitest'
import { createSession, editStructure, type BookInput } from '../src/index.js'

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
    const [n, runs] = [6000, 1600]
    const [cells, rows]: [BookInput['sheets'][0]['cells'], Record<string, { hidden: true }>] = [{ C2: { formula: `SUM(A1:A${n})` } }, {}]
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      cells[`B${r}`] = { formula: `A${r}*2` }
    }
    for (let k = 1; k <= runs; k++) rows[String(2 * k)] = { hidden: true }
    const book: BookInput = { sheets: [{ name: 'S', cells, rows, filtered: true }, { name: 'T', cells: { A1: { formula: `S!A${n}+S!A3` } } }] }
    const s = createSession(book, { execution, clock, display: false })
    // the old delete rebuilt the book once per run: 8-10 s at 800 runs of 6,000 rows, twice that at 1,600; 3 s under CI's load at 800
    const ms = timed(() => expect(s.editStructure({ op: 'deleteRows', sheet: 'S', row: 1, count: 2 * runs }).ok).toBe(true))
    expect(ms).toBeLessThan(8000)
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
    const [plain, s] = [false, true].map((filtered) => createSession({ sheets: [{ name: 'S', cells, rows, filtered }] }, { execution, clock, display: false }))
    const paste = { op: 'copyRange' as const, from: { sheet: 'S', range: 'E1' }, to: { sheet: 'S', range: `C1:C${2 * runs}` } }
    // against the same paste on the sheet unfiltered, in the same process, whatever the machine's load: the old paste
    // rebuilt the sheet once per run, 33-34 times the plain one at 400 runs of 34,000 rows, the new one 2.3 times
    const once = timed(() => expect(plain.editStructure(paste).ok).toBe(true))
    const ms = timed(() => expect(s.editStructure(paste).ok).toBe(true))
    expect(ms).toBeLessThan(10 * once)
    // the visible rows take the copy, the hidden ones keep what they had
    expect(s.getCell('S', 'C3')).toMatchObject({ formula: 'A3*3', value: 9 })
    expect(s.getCell('S', 'C2')?.formula).toBeUndefined()
    expect(s.getCell('S', `B${n}`)).toMatchObject({ formula: `A${n}*2`, value: 2 * n })
  })

  describe('a paste over runs of a filter that are no whole multiple of the copy writes the copy past its run (speed-u90)', () => {
    const book = (): BookInput => ({
      sheets: [{ name: 'S', filtered: true, rows: { '3': { hidden: true } }, cells: { A1: { value: 1 }, B1: { value: 2 }, C1: { formula: 'A1*10' }, F1: { value: 'old' }, G4: { value: 'old' }, A6: { value: 6 }, A7: { value: 7 }, E5: { value: 'below' } } }],
    })
    it.each([
      // three columns into one: each run takes one copy E:G, over the old cells beside the target
      ['A1:C1', 'E1:E5', { E1: { value: 1 }, F1: { value: 2 }, G1: { formula: 'E1*10' }, E4: { value: 1 }, F4: { value: 2 }, G4: { formula: 'E4*10' }, E5: { value: 'below' } }],
      // two rows into runs of one: the copy into E4 runs on to E5, below the target
      ['A6:A7', 'E2:E4', { F1: { value: 'old' }, G4: { value: 'old' }, E2: { value: 6 }, E3: { value: 7 }, E4: { value: 6 }, E5: { value: 7 } }],
    ])('%s onto %s, in editStructure and a session', (from, to, pasted) => {
      const op = { op: 'copyRange', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to } } as const
      const cells = { A1: { value: 1 }, B1: { value: 2 }, C1: { formula: 'A1*10' }, A6: { value: 6 }, A7: { value: 7 }, ...pasted }
      const r = editStructure(book(), op)
      expect(r.ok && r.book.sheets[0].cells).toEqual(cells)
      const s = createSession(book(), { execution, clock, display: false })
      expect(s.editStructure(op).ok).toBe(true)
      expect(s.getBook().sheets[0].cells).toEqual(cells)
    })
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
})
