/**
 * Tables under structure operations, from the review of 21.10: a copy of a whole table named as Excel would not name a
 * table, a fill continuing a series over header cells, rows inserted above a table at the sheet's edge. In each case the
 * book an operation gives, and a session's `getBook()`, computes, and the session's hash is a full computation's.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, editStructure, type BookInput, type Execution, type StructureOp } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

/** The operation in a session of each kind: its answer, and a book the session holds that computes to the session's hash. */
function inSessions(execution: Execution, book: () => BookInput, op: StructureOp, check: (r: { ok: boolean; code?: string }, after: BookInput) => void) {
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('answers as editStructure does and keeps a book that computes to its hash', async () => {
      const s = await open(book(), { execution, clock })
      const r = await s.editStructure(op)
      const after = await s.getBook()
      check(r, after)
      expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
      await s.close()
    })
  })
}

const tableBook = (name: string, names?: BookInput['names']): BookInput => ({
  ...(names ? { names } : {}),
  sheets: [{
    name: 'S',
    cells: { A1: { value: 'Qty' }, B1: { value: 'Price' }, A2: { value: 3 }, B2: { value: 1.5 }, A3: { value: 4 }, B3: { value: 2.5 }, D1: { formula: `SUM(${name}[Qty])` } },
    tables: [{ name, range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Price' }] }],
  }],
})
const copyWhole: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'A1:B3' }, to: { sheet: 'S', range: 'A10:B12' } }
const tableNames = (b: BookInput) => (b.sheets[0].tables ?? []).map((t) => t.name)

describe.each(EXECUTIONS)('a copy of a whole table (x-fix-structure-looks-1), %s', (execution) => {
  it('skips a numbered name a defined name has, at book scope or a sheet scope', () => {
    for (const names of [[{ name: 'Sales3', ref: '1' }], [{ name: 'sales3', ref: '1', scope: 'S' }, { name: 'Sales4', ref: '2' }]]) {
      const r = editStructure(tableBook('Sales', names), copyWhole)
      if (!r.ok) throw new Error(r.message)
      expect(tableNames(r.book)).toEqual(['Sales', names.length === 1 ? 'Sales4' : 'Sales5'])
      expect(compute(r.book, { execution, clock }).sheets[0].cells.D1.value).toBe(7)
    }
  })

  it('refuses a copy whose numbered name Excel would not take, as renameTable refuses it', () => {
    for (const name of ['Tab', 'Tbl', 'Rr']) {
      const r = editStructure(tableBook(name), copyWhole)
      expect(r.ok ? tableNames(r.book) : r.code, name).toBe('table_name')
    }
    // a copy of part of the table is not a table, whatever its name
    const part = editStructure(tableBook('Tab'), { op: 'copyRange', from: { sheet: 'S', range: 'A2:A3' }, to: { sheet: 'S', range: 'A10' } })
    expect(part.ok && tableNames(part.book)).toEqual(['Tab'])
  })

  inSessions(execution, () => tableBook('Tab'), copyWhole, (r, after) => {
    expect(r).toMatchObject({ ok: false, code: 'table_name' })
    expect(tableNames(after)).toEqual(['Tab'])
  })
  inSessions(execution, () => tableBook('Sales', [{ name: 'Sales3', ref: '1' }]), copyWhole, (r, after) => {
    expect(r.ok).toBe(true)
    expect(tableNames(after)).toEqual(['Sales', 'Sales4'])
  })
})

/** A table of one column at the bottom of the sheet: its header and one value above an empty last row. */
const edgeRows = (headers = true): BookInput => ({
  sheets: [{
    name: 'S',
    cells: headers ? { A1048574: { value: 'h' }, A1048575: { value: 1 } } : { A1048574: { value: 1 }, A1048575: { value: 2 } },
    tables: [{ name: 'T', range: 'A1048574:A1048576', columns: [{ name: 'h' }], ...(headers ? {} : { headers: false }) }],
  }],
})
/** A table without headers at the right edge of the sheet, its last column empty. */
const edgeColumns = (): BookInput => ({
  sheets: [{ name: 'S', cells: { XFC1: { value: 1 }, XFC2: { value: 2 } }, tables: [{ name: 'T', range: 'XFC1:XFD2', columns: [{ name: 'a' }, { name: 'b' }], headers: false }] }],
})

describe.each(EXECUTIONS)('lines inserted before a table at the sheet edge (fix-u297), %s', (execution) => {
  it.each([
    ['rows above the table', edgeRows(), { op: 'insertRows', sheet: 'S', row: 1 }, 'A1048576'],
    ['a row into the table at its last row', edgeRows(), { op: 'insertRows', sheet: 'S', row: 1048576 }, 'A1048576'],
    ['rows above a table without headers', edgeRows(false), { op: 'insertRows', sheet: 'S', row: 1 }, 'A1048576'],
    ['columns before a table without headers', edgeColumns(), { op: 'insertColumns', sheet: 'S', column: 'A' }, 'XFD2'],
  ] as [string, BookInput, StructureOp, string][])('refuses %s with off_sheet, naming the table\'s last cell', (_, book, op, cell) => {
    compute(book, { execution, clock })
    const r = editStructure(book, op)
    expect(r.ok ? r.book.sheets[0].tables : r.code).toBe('off_sheet')
    expect(!r.ok && r.message).toContain(`S!${cell}`)
  })

  it('takes rows inserted above a table that ends before the edge', () => {
    const book = edgeRows()
    book.sheets[0].tables![0].range = 'A1048574:A1048575'
    const r = editStructure(book, { op: 'insertRows', sheet: 'S', row: 1 })
    if (!r.ok) throw new Error(r.message)
    expect(r.book.sheets[0].tables![0].range).toBe('A1048575:A1048576')
    compute(r.book, { execution, clock })
  })

  inSessions(execution, edgeRows, { op: 'insertRows', sheet: 'S', row: 1 }, (r, after) => {
    expect(r).toMatchObject({ ok: false, code: 'off_sheet' })
    expect(after.sheets[0].tables![0].range).toBe('A1048574:A1048576')
  })
})
