/**
 * Excel's Sort (`sortRange`, and the sort `applyFilter` reapplies): the rows shown reordered, ties in their order, blanks
 * last either way, hidden rows in place; a cell's formula moved as a copy, references from elsewhere left; the sort kept
 * where Excel keeps it; refusals; in a session with undo. The sorts of the books Excel made are fixed points of it.
 */
import { readFileSync, readdirSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { createSession, editStructure, type BookInput, type FilterSort, type SheetInput } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-11-10T12:00:00Z', seed: 1 }

/** Data: names in A2:A7, amounts in B, C = B*10, and E1 summing B2 outside the range. */
const book = (data: Partial<SheetInput> = {}): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 'Name' }, B1: { value: 'Amount' }, E1: { formula: 'B2' },
        A2: { value: 'b' }, B2: { value: 3 }, A3: { value: 'B' }, B3: { value: 'x' }, A4: { value: 'a' }, B4: { value: 1 },
        A5: { value: 'c' }, A6: { value: 'A' }, B6: { value: true }, A7: { value: 'd' }, B7: { value: 1 },
        ...Object.fromEntries([2, 3, 4, 5, 6, 7].map((r) => [`C${r}`, { formula: `B${r}*10` }])),
      },
      ...data,
    },
  ],
})

const sorted = (start: BookInput, sort: FilterSort) => {
  const r = editStructure(start, { op: 'sortRange', sheet: 'Data', sort })
  if (!r.ok) throw new Error(r.message)
  return r.book.sheets[0]
}

const column = (s: SheetInput, c: string) => [2, 3, 4, 5, 6, 7].map((r) => s.cells[`${c}${r}`]?.value ?? s.cells[`${c}${r}`]?.formula ?? null)

describe('sortRange', () => {
  it('numbers, text, booleans, then blanks, ties in their order; descending keeps blanks last', () => {
    const s = sorted(book(), { range: 'A2:C7', conditions: [{ range: 'B2:B7' }] })
    expect(column(s, 'A')).toEqual(['a', 'd', 'b', 'B', 'A', 'c'])
    expect(column(sorted(book(), { range: 'A2:C7', conditions: [{ range: 'B2:B7', descending: true }] }), 'A')).toEqual(['A', 'B', 'b', 'a', 'd', 'c'])
  })

  it('moves formulas as copies, their own row with them; a reference from outside keeps its address', () => {
    const s = sorted(book(), { range: 'A2:C7', conditions: [{ range: 'B2:B7' }] })
    expect(column(s, 'C')).toEqual(['B2*10', 'B3*10', 'B4*10', 'B5*10', 'B6*10', 'B7*10'])
    expect(s.cells.E1).toEqual({ formula: 'B2' })
    expect(s.sort).toEqual({ range: 'A2:C7', conditions: [{ range: 'B2:B7' }] })
  })

  it('case-insensitive by default, lower case first with caseSensitive; a custom list first', () => {
    expect(column(sorted(book(), { range: 'A2:C7', conditions: [{ range: 'A2:A7' }] }), 'A')).toEqual(['a', 'A', 'b', 'B', 'c', 'd'])
    expect(column(sorted(book(), { range: 'A2:C7', caseSensitive: true, conditions: [{ range: 'A2:A7' }] }), 'A')).toEqual(['a', 'A', 'b', 'B', 'c', 'd'])
    const upper = book()
    upper.sheets[0].cells.A4 = { value: 'A' }
    upper.sheets[0].cells.A6 = { value: 'a' }
    expect(column(sorted(upper, { range: 'A2:C7', caseSensitive: true, conditions: [{ range: 'A2:A7' }] }), 'A')).toEqual(['a', 'A', 'b', 'B', 'c', 'd'])
    expect(column(sorted(book(), { range: 'A2:C7', conditions: [{ range: 'A2:A7', customList: 'd,c' }] }), 'A')).toEqual(['d', 'c', 'a', 'A', 'b', 'B'])
  })

  it('leaves hidden rows where they are', () => {
    const s = sorted(book({ rows: { 4: { hidden: true } } }), { range: 'A2:C7', conditions: [{ range: 'A2:A7' }] })
    expect(column(s, 'A')).toEqual(['A', 'b', 'a', 'B', 'c', 'd'])
  })

  it('a cell color on top', () => {
    const start = book()
    start.styles = [{ fill: { color: '#FF0000' } }]
    start.sheets[0].cells.A5.style = 0
    expect(column(sorted(start, { range: 'A2:C7', conditions: [{ range: 'A2:A7', by: 'cellColor', style: { fill: { color: '#FF0000' } } }] }), 'A')).toEqual(['c', 'b', 'B', 'a', 'A', 'd'])
  })

  it("keeps the sort as the sheet's over an AutoFilter's or a table's data too, as Worksheet.Sort does", () => {
    const sort = { range: 'A2:C7', conditions: [{ range: 'A2:A7' }] }
    const af = sorted(book({ autoFilter: { range: 'A1:C7' } }), sort)
    expect([af.autoFilter, af.sort]).toEqual([{ range: 'A1:C7' }, sort])
    const start = book({ tables: [{ name: 'T', range: 'A1:C7', columns: [{ name: 'Name' }, { name: 'Amount' }, { name: 'Column1' }] }] })
    start.sheets[0].cells.C1 = { value: 'Column1' }
    const t = sorted(start, sort)
    expect([t.tables?.[0].filter, t.sort]).toEqual([undefined, sort])
  })

  it('refuses merged areas, part of an array, part of a table; a condition outside is an input error', () => {
    const sort = { range: 'A2:C7', conditions: [{ range: 'A2:A7' }] }
    expect(editStructure(book({ merges: ['A3:B3'] }), { op: 'sortRange', sheet: 'Data', sort })).toMatchObject({ ok: false, code: 'merged' })
    const array = book()
    array.sheets[0].cells.C3 = { formula: 'B3:B4*10', array: 'C3:C4' }
    delete array.sheets[0].cells.C4
    expect(editStructure(array, { op: 'sortRange', sheet: 'Data', sort })).toMatchObject({ ok: false, code: 'array' })
    const table = book({ tables: [{ name: 'T', range: 'B1:C4', columns: [{ name: 'Amount' }, { name: 'Column1' }] }] })
    table.sheets[0].cells.C1 = { value: 'Column1' }
    expect(editStructure(table, { op: 'sortRange', sheet: 'Data', sort })).toMatchObject({ ok: false, code: 'table', message: 'the range cannot be sorted over part of the table T' })
    expect(() => editStructure(book(), { op: 'sortRange', sheet: 'Data', sort: { range: 'A2:C7', conditions: [{ range: 'E2:E7' }] } })).toThrow(/outside the sorted range/)
  })

  it('refuses a range that meets a spilled array but whole in one of its lines, as Excel does; one whole in a line moves with it', () => {
    const spill: BookInput = { sheets: [{ name: 'S', cells: { G1: { formula: 'SEQUENCE(3,,3,-1)' }, H1: { value: 1 }, H2: { value: 2 }, H3: { value: 3 } } }] }
    // the condition is the range's first column
    const first = (range: string) => {
      const [a, b] = range.split(':')
      return `${a}:${a.replace(/\d/g, '')}${b.replace(/[A-Z]/g, '')}`
    }
    for (const range of ['G1:G3', 'G1:G5', 'G2:G3', 'G1:H3'])
      expect(editStructure(spill, { op: 'sortRange', sheet: 'S', sort: { range, conditions: [{ range: first(range) }] } }), range).toMatchObject({ ok: false, code: 'array' })
    const s = createSession(spill)
    try {
      expect(s.editStructure({ op: 'sortRange', sheet: 'S', sort: { range: 'G1:G3', conditions: [{ range: 'G1:G3' }] } })).toMatchObject({ ok: false, code: 'array' })
      expect(s.getBook().sheets[0].cells.G1).toEqual({ formula: 'SEQUENCE(3,,3,-1)' })
      expect(s.getCells('S', 'G1:G3').G3.value).toBe(1)
    } finally {
      s.close()
    }
    // Excel for Windows: a spill in a row sorts with its row, one in a column with its column, a legacy array alike
    const rest = { G2: { value: 5 }, H2: { value: 6 }, I2: { value: 7 }, G3: { value: 9 }, H3: { value: 8 }, I3: { value: 7 } }
    const values = (book: BookInput, range: string) => {
      const out = createSession(book)
      try {
        return Object.values(out.getCells('S', range)).map((c) => c.value)
      } finally {
        out.close()
      }
    }
    for (const g1 of [{ formula: 'SEQUENCE(1,3)' }, { formula: 'COLUMN(A1:C1)', array: 'G1:I1' }]) {
      const r = editStructure({ sheets: [{ name: 'S', cells: { G1: g1, ...rest } }] }, { op: 'sortRange', sheet: 'S', sort: { range: 'G1:I3', conditions: [{ range: 'G1:G3', descending: true }] } })
      expect(r.ok, g1.formula).toBe(true)
      if (r.ok) expect(values(r.book, 'G1:I3')).toEqual([9, 8, 7, 5, 6, 7, 1, 2, 3])
    }
    const columns = editStructure({ sheets: [{ name: 'S', cells: { G1: { formula: 'SEQUENCE(3)' }, H1: { value: 5 }, H2: { value: 6 }, H3: { value: 7 }, I1: { value: 9 }, I2: { value: 8 }, I3: { value: 7 } } }] }, { op: 'sortRange', sheet: 'S', sort: { range: 'G1:I3', byColumns: true, conditions: [{ range: 'G1:I1', descending: true }] } })
    expect(columns.ok).toBe(true)
    if (columns.ok) expect(values(columns.book, 'G1:I3')).toEqual([9, 5, 1, 8, 6, 2, 7, 7, 3])
  })
})

describe('applyFilter with a sort', () => {
  it('sorts the rows it shows, the hidden ones in place', () => {
    const r = editStructure(book({ autoFilter: { range: 'A1:C7' } }), {
      op: 'applyFilter',
      sheet: 'Data',
      criteria: { columns: [{ column: 0, filter: { type: 'custom', conditions: [{ operator: 'notEqual', value: 'c' }] } }], sort: { range: 'A2:C7', conditions: [{ range: 'A2:A7', descending: true }] } },
    })
    if (!r.ok) throw new Error(r.message)
    expect(column(r.book.sheets[0], 'A')).toEqual(['d', 'b', 'B', 'c', 'a', 'A'])
    expect(r.book.sheets[0].rows).toEqual({ 5: { hidden: true } })
  })
})

describe.each(KINDS)('a sort in %s', (kind) => {
  it('changes the values of the cells it sorts, keeps every address, undo and redo', async () => {
    const s = await (await opener(kind))(book(), { clock })
    const r = await s.editStructure({ op: 'sortRange', sheet: 'Data', sort: { range: 'A2:C7', conditions: [{ range: 'B2:B7' }] } })
    if (!r.ok) throw new Error(r.message)
    expect(r.mapCell({ sheet: 'Data', cell: 'B4' })).toEqual({ sheet: 'Data', cell: 'B4' })
    expect((await s.getCell('Data', 'C2'))?.value).toBe(10)
    expect((await s.getCell('Data', 'E1'))?.value).toBe(1)
    await s.undo()
    expect((await s.getCell('Data', 'A2'))?.value).toBe('b')
    expect((await s.getCell('Data', 'E1'))?.value).toBe(3)
    await s.redo()
    expect((await s.getCell('Data', 'A2'))?.value).toBe('a')
    await s.close()
  })
})

/** Books whose rows Excel did not sort again after an edit that broke their order (a row inserted, the range moved). */
const STALE = new Set(['range-sort-row-inside', 'sheet-row-inside', 'table-row-inside', 'sheet-moved', 'table-moved'])
const made = new URL('./excel/win/filter-criteria-made/', import.meta.url)
const fixed: [string, string, FilterSort, BookInput, number][] = []
for (const file of readdirSync(made).filter((n) => n.endsWith('.xlsx')).sort()) {
  const name = file.slice(0, -5)
  if (STALE.has(name)) continue
  const { book: b } = await readXlsx(readFileSync(new URL(file, made)))
  b.sheets.forEach((s, i) => {
    for (const sort of [s.sort, s.autoFilter?.criteria?.sort, ...(s.tables ?? []).map((t) => (t.filter ? t.filter.sort : undefined))]) if (sort) fixed.push([name, s.name, sort, b, i])
  })
}

it('finds the sorts of the books Excel made', () => {
  expect(fixed.length).toBe(53)
})

it.each(fixed)('%s %s: the order Excel sorted is a fixed point', (_, sheet, sort, b, i) => {
  const s = createSession(b, { clock })
  const before = s.getBook().sheets[i].cells
  expect(s.editStructure({ op: 'sortRange', sheet, sort }).ok).toBe(true)
  expect(s.getBook().sheets[i].cells).toEqual(before)
})
