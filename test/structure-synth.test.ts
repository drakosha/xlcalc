/**
 * Structure operations on the synthetic book (340 data rows): inserting rows or columns anywhere changes
 * no value (the book has no function that counts blanks or positions), and deleting what was inserted
 * gives back the same book and the same hash.
 */
import { describe, expect, it } from 'vitest'
import { compute, editStructure, type BookInput, type CellRef, type StructureOp } from '../src/index.js'
import { S, synthBook } from '../src/synth.js'
import { cellIssues } from './diagnostics.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const { book } = synthBook({ seed: 1, rows: 340 })
const before = compute(book, { clock })

function apply(b: BookInput, op: StructureOp) {
  const r = editStructure(b, op)
  if (!r.ok) throw new Error(r.message)
  return r
}

/** Every value of the old book is at its new place. */
function sameValues(op: StructureOp) {
  const r = apply(book, op)
  const after = compute(r.book, { clock })
  const off: string[] = []
  before.sheets.forEach((s) => {
    const now = after.sheets.find((x) => x.name === s.name)!
    for (const [cell, c] of Object.entries(s.cells)) {
      const to = r.mapCell({ sheet: s.name, cell }) as CellRef
      const v = now.cells[to.cell]?.value ?? null
      if (JSON.stringify(v) !== JSON.stringify(c.value)) off.push(`${s.name}!${cell} → ${to.cell}: ${JSON.stringify(c.value)} → ${JSON.stringify(v)}`)
    }
  })
  expect(off.slice(0, 5)).toEqual([])
  expect(cellIssues(after).length).toBe(cellIssues(before).length)
  return r.book
}

const inserts: StructureOp[] = [
  { op: 'insertRows', sheet: S.data, row: 2, count: 3 },
  { op: 'insertRows', sheet: S.data, row: 170, count: 1 },
  { op: 'insertRows', sheet: S.data, row: 342, count: 2 },
  { op: 'insertColumns', sheet: S.data, column: 'A', count: 2 },
  { op: 'insertColumns', sheet: S.data, column: 'E', count: 1 },
  { op: 'insertRows', sheet: S.items, row: 5, count: 2 },
  { op: 'insertColumns', sheet: S.items, column: 'D', count: 1 },
  { op: 'insertRows', sheet: S.params, row: 1, count: 4 },
  { op: 'insertColumns', sheet: S.params, column: 'B', count: 1 },
  { op: 'insertRows', sheet: S.months, row: 3, count: 1 },
  { op: 'insertRows', sheet: S.months, row: 8, count: 3 },
  { op: 'insertColumns', sheet: S.months, column: 'C', count: 2 },
  { op: 'insertColumns', sheet: S.projects, column: 'P', count: 1 },
  { op: 'insertRows', sheet: S.cats, row: 2, count: 1 },
  { op: 'insertColumns', sheet: S.dirs, column: 'A', count: 1 },
]

const undo = (op: StructureOp): StructureOp => (op.op === 'insertRows' ? { ...op, op: 'deleteRows' } : op.op === 'insertColumns' ? { ...op, op: 'deleteColumns' } : op)

describe('structure operations on the synthetic book', () => {
  for (const op of inserts) {
    const at = 'row' in op ? op.row : (op as { column: string }).column
    it(`${op.op} ${op.op.includes('Rows') ? 'rows' : 'columns'} at ${at} of ${(op as { sheet: string }).sheet}: values unchanged, deleting them gives the book back`, () => {
      const next = sameValues(op)
      const back = apply(next, undo(op)).book
      expect(back).toEqual(book)
      expect(compute(back, { clock }).hash).toBe(before.hash)
    })
  }

  it('renaming every sheet and back gives the book back; values follow the names', () => {
    let b = book
    for (const s of book.sheets) b = apply(b, { op: 'renameSheet', sheet: s.name, name: `${s.name} 2` }).book
    const values = (cells: Record<string, { value: unknown }>) => Object.fromEntries(Object.entries(cells).map(([a, c]) => [a, c.value]))
    const renamed = compute(b, { clock })
    before.sheets.forEach((s, i) => expect(values(renamed.sheets[i].cells)).toEqual(values(s.cells)))
    for (const s of book.sheets) b = apply(b, { op: 'renameSheet', sheet: `${s.name} 2`, name: s.name }).book
    expect(b).toEqual(book)
  })

  it('moving a report block to an empty place and to a new sheet keeps every value', () => {
    const block = { sheet: S.months, range: 'A1:N50' }
    sameValues({ op: 'moveRange', from: block, to: { sheet: S.months, cell: 'A101' } })
    const withSheet: BookInput = { ...book, sheets: [...book.sheets, { name: 'Пустой', cells: {} }] }
    const r = apply(withSheet, { op: 'moveRange', from: block, to: { sheet: 'Пустой', cell: 'C3' } })
    const after = compute(r.book, { clock })
    for (const [cell, c] of Object.entries(before.sheets.find((s) => s.name === S.months)!.cells)) {
      const to = r.mapCell({ sheet: S.months, cell })!
      expect(after.sheets.find((s) => s.name === to.sheet)!.cells[to.cell].value).toEqual(c.value)
    }
    expect(after.sheets.find((s) => s.name === S.params)!.cells.B12.value).toEqual(before.sheets.find((s) => s.name === S.params)!.cells.B12.value)
  })
})
