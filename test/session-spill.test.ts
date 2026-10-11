/**
 * Sessions over spilled arrays: after each of 200 seeded edits (`session-spill-edits.ts`) the session's hash
 * equals a full computation's, and the changes it reported agree with its result; a value typed into a spill
 * range stops the array, clearing it lets it spill again; merged areas through `setLayout`; spill fields in
 * changes and `explain`.
 */
import { describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput, type CellChange, type Session } from '../src/index.js'
import { apply } from './session-edits.js'
import { spillBook, spillEdits } from './session-spill-edits.js'
import { shownValue } from './spill-shown.js'
import { toVal } from './excel-cases.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 5 }

/** Every change agrees with the session's result: its value as shown, its spill fields (a cell of a fill names its formula). */
function agrees(s: Session, changes: CellChange[]) {
  const r = s.result()
  for (const ch of changes) {
    const cells = r.sheets.find((x) => x.name === ch.sheet)!.cells
    const cell = cells[ch.cell]
    expect(shownValue(r, ch.sheet, ch.cell), `${ch.sheet}!${ch.cell}`).toEqual(ch.value === null ? null : toVal(ch.value))
    const fill = !cell && ch.spillOf ? { spillOf: ch.spillOf } : {}
    expect({ spill: ch.spill, spillFill: ch.spillFill, spillOf: ch.spillOf }).toEqual({ spill: cell?.spill, spillFill: cell?.spillFill, spillOf: cell?.spillOf, ...fill })
    if (!cell && ch.spillOf) expect(cells[ch.spillOf].spillFill).toBeDefined()
  }
}

/** Changes in the order of their cells' addresses (the session reports them in the order it met them). */
const sorted = (r: { changes: CellChange[] }) => [...r.changes].sort((a, b) => a.sheet.localeCompare(b.sheet) || a.cell.length - b.cell.length || a.cell.localeCompare(b.cell))

const S = (cell: string, value: number | string | null) => ({ sheet: 'S', cell, input: value === null ? null : { value } })

describe('sessions with spilled arrays', () => {
  it('200 edits: after each one the session hash equals the full computation hash', () => {
    const s = openSession(spillBook(), { clock })
    let book = spillBook()
    for (const batch of spillEdits(11, 200)) {
      const { changes } = s.setCells(batch)
      book = apply(book, batch)
      expect(s.result().hash).toBe(compute(book, { clock }).hash)
      agrees(s, changes)
    }
  }, 120_000)

  it('a value typed into a spill range stops the array; clearing it lets the array spill again', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(3)' }, B1: { formula: 'SUM(A1:A3)' } } }] }, { clock, display: false })
    expect(sorted(s.setCells([S('A3', 'x')]))).toEqual([
      { sheet: 'S', cell: 'A1', value: { error: 'SPILL' } },
      { sheet: 'S', cell: 'A2', value: null },
      { sheet: 'S', cell: 'A3', value: 'x' },
      { sheet: 'S', cell: 'B1', value: { error: 'SPILL' } },
    ])
    expect(s.result().issues[0]).toMatchObject({ cell: 'A1', reason: "the array of 3×1 values cannot spill: 'S'!A3 is not empty" })
    expect(sorted(s.setCells([S('A3', null)]))).toEqual([
      { sheet: 'S', cell: 'A1', value: 1, spill: 'A1:A3' },
      { sheet: 'S', cell: 'A2', value: 2, spillOf: 'A1' },
      { sheet: 'S', cell: 'A3', value: 3, spillOf: 'A1' },
      { sheet: 'S', cell: 'B1', value: 6 },
    ])
  })

  it('a merged area added over a spill range stops the array, removed lets it spill', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(1,3)' } } }] }, { clock })
    expect(sorted(s.setLayout([{ sheet: 'S', merges: ['C1:C2'] }])).map((c) => [c.cell, c.value])).toEqual([
      ['A1', { error: 'SPILL' }],
      ['B1', null],
      ['C1', null],
    ])
    expect(s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 20 } }]).changes).toEqual([])
    expect(sorted(s.setLayout([{ sheet: 'S', merges: [] }])).map((c) => [c.cell, c.value])).toEqual([
      ['A1', 1],
      ['B1', 2],
      ['C1', 3],
    ])
  })

  it('an array whose size changes moves its readers; E2# follows it', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { B1: { value: 2 }, A1: { formula: 'SEQUENCE(B1)' }, C1: { formula: 'SUM(A1#)' }, C2: { formula: 'A3+1' } } }] }
    const s = openSession(book, { clock })
    expect(sorted(s.setCells([S('B1', 4)])).map((c) => [c.cell, c.value, c.spill])).toEqual([
      ['A1', 1, 'A1:A4'],
      ['A3', 3, undefined],
      ['A4', 4, undefined],
      ['B1', 4, undefined],
      ['C1', 10, undefined],
      ['C2', 4, undefined],
    ])
    expect(s.result().hash).toBe(compute(apply(book, [S('B1', 4)]), { clock }).hash)
  })

  it('a whole column: the cells past the data are one fill, reported with the formula', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A:A*2' }, C1: { formula: 'COUNT(B:B)' } } }] }, { clock, display: false })
    const b1 = s.result().sheets[0].cells.B1
    expect(b1).toEqual({ value: 2, formula: 'A:A*2', spill: 'B1:B1048576', spillFill: 0 })
    expect(Object.keys(s.result().sheets[0].cells)).toEqual(['A1', 'B1', 'C1', 'A2', 'B2'])
    expect(s.result().sheets[0].cells.C1.value).toBe(1048576)
    expect(sorted(s.setCells([S('B9', 'x')])).map((c) => [c.cell, c.value, c.spill])).toEqual([
      ['B1', { error: 'SPILL' }, undefined],
      ['B2', null, undefined],
      ['B9', 'x', undefined],
      ['C1', 0, undefined],
    ])
  })

  it('a cycle through a spill range is #CYCLE!, and breaking it computes again', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { B1: { value: 2 }, A1: { formula: 'SEQUENCE(B1)' } } }] }
    const s = openSession(book, { clock, display: false })
    const cyc = [{ sheet: 'S', cell: 'B1', input: { formula: 'A2+1' } }]
    s.setCells(cyc)
    expect(s.result().hash).toBe(compute(apply(book, cyc), { clock }).hash)
    s.setCells([S('B1', 3)])
    expect(s.result().sheets[0].cells.A1).toEqual({ value: 1, formula: 'SEQUENCE(B1)', spill: 'A1:A3' })
  })

  it('explain: a spilled cell leads to its formula', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { value: 4 }, B1: { formula: 'SEQUENCE(2)*A1' } } }] }, { clock })
    const e = s.explain('S', 'B2')
    expect(e.root).toMatchObject({ ref: "'S'!B2", value: 8, spillOf: "'S'!B1", children: [{ ref: "'S'!B1", formula: 'SEQUENCE(2)*A1' }] })
    expect(e.text).toBe(["'S'!B2 = 8  ← spilled from 'S'!B1", "  'S'!B1 = 4  ← SEQUENCE(2)*A1", "    'S'!A1 = 4"].join('\n'))
  })
})
