/**
 * A filter applied (`applyFilter`, `clearFilter`): the rows of a sheet's AutoFilter or a table shown and hidden by its
 * criteria as Excel's Reapply shows and hides them, the criteria kept, the filter mode set; a last row of SUBTOTAL and a
 * table's totals row left out; in every kind of session, with undo and redo. `filter-apply-excel.test.ts` compares the
 * rows with the books Excel filtered.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { createSession, editStructure, InputError, type BookInput, type FilterCriteria, type SheetInput, type StructureOp } from '../src/index.js'
import { lines } from '../src/layout/spans.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-11-10T12:00:00Z', seed: 1 }

/** Data: Region and Amount over A1:B7 with an AutoFilter, a SUBTOTAL of the amounts shown in D1. */
const book = (data: Partial<SheetInput> = {}): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 'Region' }, B1: { value: 'Amount' }, D1: { formula: 'SUBTOTAL(109,B2:B7)' },
        A2: { value: 'North' }, B2: { value: 1 }, A3: { value: 'south ' }, B3: { value: 2 }, A4: { value: 'East' }, B4: { value: 3 },
        A5: { value: 'South' }, B5: { value: 4 }, A6: { value: 'West' }, B6: { value: 5 }, A7: { value: 'North' }, B7: { value: 6 },
      },
      autoFilter: { range: 'A1:B7' },
      ...data,
    },
  ],
})

const south: FilterCriteria = { columns: [{ column: 0, filter: { type: 'values', values: ['South'] } }] }

/** The hidden rows of a sheet, by number. */
const hidden = (s: Pick<SheetInput, 'rows'>) => lines(s.rows, 'r').filter(([, , v]) => v.hidden).flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => a + i + 1))

const applied = (start: BookInput, op: StructureOp) => {
  const r = editStructure(start, op)
  if (!r.ok) throw new Error(r.message)
  return r.book.sheets[0]
}

describe('applyFilter', () => {
  it('sets the criteria and shows the rows they keep, one hidden by hand too, by the text shown in any case and without its spaces', () => {
    const s = applied(book({ rows: { 5: { hidden: true, height: 20 } } }), { op: 'applyFilter', sheet: 'Data', criteria: south })
    expect(hidden(s)).toEqual([2, 4, 6, 7])
    expect(s.rows?.['5']).toEqual({ height: 20 })
    expect(s.filtered).toBe(true)
    expect(s.autoFilter).toEqual({ range: 'A1:B7', criteria: south })
  })

  it('applies the criteria held again (Reapply), over the range the AutoFilter acts on, which it takes', () => {
    const start = book({ autoFilter: { range: 'A1:B5', criteria: south } })
    start.sheets[0].cells.A8 = { value: 'South' }
    const s = applied(start, { op: 'applyFilter', sheet: 'Data' })
    expect(hidden(s)).toEqual([2, 4, 6, 7])
    expect(s.autoFilter?.range).toBe('A1:B8')
  })

  it('neither filters nor saves a last row of SUBTOTAL, as Excel', () => {
    const start = book()
    start.sheets[0].cells.B8 = { formula: 'SUBTOTAL(9,B2:B7)' }
    const s = applied(start, { op: 'applyFilter', sheet: 'Data', criteria: south })
    expect(hidden(s)).toEqual([2, 4, 6, 7])
    expect(s.autoFilter?.range).toBe('A1:B7')
  })

  it('combines the columns: each must keep a row', () => {
    const criteria: FilterCriteria = {
      columns: [
        { column: 0, filter: { type: 'custom', conditions: [{ operator: 'notEqual', value: 'East' }] } },
        { column: 1, filter: { type: 'custom', and: true, conditions: [{ operator: 'greaterThan', value: '1' }, { operator: 'lessThanOrEqual', value: '5' }] } },
      ],
    }
    expect(hidden(applied(book(), { op: 'applyFilter', sheet: 'Data', criteria }))).toEqual([2, 4, 7])
  })

  it("filters a table's data rows, its totals row left as it is", () => {
    const start = book({ autoFilter: undefined, tables: [{ name: 'T', range: 'A1:B8', totals: true, columns: [{ name: 'Region' }, { name: 'Amount', totals: { function: 'sum' } }] }] })
    start.sheets[0].cells.B8 = { formula: 'SUBTOTAL(109,T[Amount])' }
    const s = applied(start, { op: 'applyFilter', table: 't', criteria: south })
    expect(hidden(s)).toEqual([2, 4, 6, 7])
    expect(s.tables?.[0].filter).toEqual(south)
    expect(s.filtered).toBe(true)
  })

  it('a sheet without an AutoFilter, a table without buttons, a column outside the filter are input errors', () => {
    const noFilter = book({ autoFilter: undefined })
    expect(() => editStructure(noFilter, { op: 'applyFilter', sheet: 'Data' })).toThrow(InputError)
    const noButtons = book({ autoFilter: undefined, tables: [{ name: 'T', range: 'A1:B7', columns: [{ name: 'Region' }, { name: 'Amount' }], filter: false }] })
    expect(() => editStructure(noButtons, { op: 'applyFilter', table: 'T' })).toThrow(/no filter buttons/)
    expect(() => editStructure(book(), { op: 'applyFilter', sheet: 'Data', criteria: { columns: [{ column: 2, filter: { type: 'values', values: ['x'] } }] } })).toThrow(InputError)
    expect(() => editStructure(book(), { op: 'applyFilter', sheet: 'Data', table: 'T' } as never)).toThrow(/one of sheet and table/)
  })
})

describe('the kinds of filter', () => {
  const by = (filter: NonNullable<FilterCriteria['columns']>[number]['filter'], start = book(), column = 1) => {
    const r = editStructure(start, { op: 'applyFilter', sheet: 'Data', criteria: { columns: [{ column, filter }] } })
    if (!r.ok) throw new Error(r.message)
    return r.book.sheets[0]
  }

  it('top and bottom items and percents over the whole column, ties kept, the bound stored', () => {
    const start = book()
    start.sheets[0].cells.B3 = { value: 5 }
    const s = by({ type: 'top', value: 2 }, start)
    expect(hidden(s)).toEqual([2, 4, 5])
    expect(s.autoFilter?.criteria?.columns?.[0].filter).toEqual({ type: 'top', value: 2, filterValue: 5 })
    expect(hidden(by({ type: 'top', bottom: true, percent: true, value: 40 }))).toEqual([4, 5, 6, 7])
  })

  it('above and below the average, strictly, the average stored', () => {
    const s = by({ type: 'dynamic', kind: 'aboveAverage' })
    expect(hidden(s)).toEqual([2, 3, 4])
    expect(s.autoFilter?.criteria?.columns?.[0].filter).toEqual({ type: 'dynamic', kind: 'aboveAverage', value: 3.5 })
  })

  it('refuses a top or average filter over an error, in the locale', () => {
    const start = book()
    start.sheets[0].cells.B4 = { formula: 'NA()' }
    const r = editStructure(start, { op: 'applyFilter', sheet: 'Data', criteria: { columns: [{ column: 1, filter: { type: 'top', value: 1 } }] } })
    expect(r).toEqual({ ok: false, code: 'filter', message: 'a top or average filter cannot be applied over the error in Data!B4' })
  })

  it('relative dates from the day of the calculation, weeks from Sunday, by numbers shown as dates', () => {
    // 2026-11-10 is a Tuesday: the week is 2026-11-08 to 2026-11-14
    const start = book()
    const cells = start.sheets[0].cells
    for (const [row, date] of [[2, '2026-11-07'], [3, '2026-11-08'], [4, '2026-11-14'], [5, '2026-11-15'], [6, '2026-11-10']] as const) cells[`B${row}`] = { value: { date }, numberFormat: 'yyyy-mm-dd' }
    cells.B7 = { value: { date: '2026-11-10' } }
    const s = createSession(start, { clock })
    expect(s.editStructure({ op: 'applyFilter', sheet: 'Data', criteria: { columns: [{ column: 1, filter: { type: 'dynamic', kind: 'thisWeek' } }] } }).ok).toBe(true)
    const sheet = s.getBook().sheets[0]
    expect(hidden(sheet)).toEqual([2, 5, 7])
    expect(sheet.autoFilter?.criteria?.columns?.[0].filter).toEqual({ type: 'dynamic', kind: 'thisWeek', value: 46334, maxValue: 46341 })
  })

  it("a fill color, a conditional format's first; a font color of the cell or of one run of its rich text; an icon", () => {
    const start = book({
      conditionalFormats: [
        { type: 'cellIs', ranges: ['B2:B7'], operator: 'equal', formulas: ['3'], style: { fill: { color: '#FF0000' } } },
        { type: 'iconSet', ranges: ['B2:B7'], icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] },
      ],
    })
    start.styles = [{ fill: { color: '#FF0000' } }, { font: { color: '#0000FF' } }]
    const cells = start.sheets[0].cells
    cells.B5.style = 0
    cells.A3 = { value: 'south', richText: [{ start: 0 }, { start: 2, font: { color: '#0000FF' } }] }
    cells.A6.style = 1
    expect(hidden(by({ type: 'color', style: { fill: { color: '#FF0000' } } }, start))).toEqual([2, 3, 6, 7])
    expect(hidden(by({ type: 'color', style: {} }, start))).toEqual([4, 5])
    expect(hidden(by({ type: 'color', font: true, style: { font: { color: '#0000FF' } } }, start, 0))).toEqual([2, 4, 5, 7])
    expect(hidden(by({ type: 'icon', set: '3Arrows', icon: 2 }, start))).toEqual([2, 3, 4, 5])
  })
})

describe('the criteria as Excel saves them after applying them', () => {
  const after = (filter: NonNullable<FilterCriteria['columns']>[number]['filter'], column = 0) => {
    const r = editStructure(book(), { op: 'applyFilter', sheet: 'Data', criteria: { columns: [{ column, filter }] } })
    if (!r.ok) throw new Error(r.message)
    return r.book.sheets[0].autoFilter?.criteria?.columns?.[0].filter
  }

  it('values in order of their text, dates in time order', () => {
    expect(after({ type: 'values', values: ['West', '#N/A', '25%', 'east'], dates: [{ level: 'day', year: 2026, month: 11, day: 10 }, { level: 'month', year: 2026, month: 10 }] })).toEqual({
      type: 'values',
      values: ['#N/A', '25%', 'east', 'West'],
      dates: [{ level: 'month', year: 2026, month: 10 }, { level: 'day', year: 2026, month: 11, day: 10 }],
    })
  })

  it('an operand that reads as a number or a date as that number; equality alone as values', () => {
    expect(after({ type: 'custom', conditions: [{ operator: 'greaterThan', value: '2026-11-01' }, { operator: 'lessThan', value: '50%' }] }, 1)).toEqual({ type: 'custom', conditions: [{ operator: 'greaterThan', value: '46327' }, { operator: 'lessThan', value: '0.5' }] })
    expect(after({ type: 'custom', conditions: [{ value: 'West' }, { operator: 'equal', value: '7' }] })).toEqual({ type: 'values', values: ['7', 'West'] })
    expect(after({ type: 'custom', conditions: [{ value: 'W*' }] })).toEqual({ type: 'custom', conditions: [{ value: 'W*' }] })
  })
})

describe('clearFilter', () => {
  const both: FilterCriteria = { columns: [{ column: 0, hiddenButton: true, filter: { type: 'values', values: ['South', 'North'] } }, { column: 1, filter: { type: 'custom', conditions: [{ operator: 'greaterThan', value: '4' }] } }], sort: { range: 'A2:B7', conditions: [{ range: 'B2:B7' }] } }

  it("removes every column's condition and the sort, keeps the buttons, and shows the rows", () => {
    const start = applied(book(), { op: 'applyFilter', sheet: 'Data', criteria: both })
    expect(hidden(start)).toEqual([2, 3, 4, 5, 6])
    const s = applied({ sheets: [start] }, { op: 'clearFilter', sheet: 'Data' })
    expect(hidden(s)).toEqual([])
    expect(s.autoFilter).toEqual({ range: 'A1:B7', criteria: { columns: [{ column: 0, hiddenButton: true }] } })
    expect('filtered' in s).toBe(false)
  })

  it("removes one column's condition and applies the others and the sort again", () => {
    const start = applied(book(), { op: 'applyFilter', sheet: 'Data', criteria: both })
    const s = applied({ sheets: [start] }, { op: 'clearFilter', sheet: 'Data', column: 1 })
    expect(hidden(s)).toEqual([4, 6])
    expect(s.filtered).toBe(true)
    expect(s.autoFilter?.criteria?.sort).toEqual(both.sort)
    expect([2, 3, 4, 5, 6, 7].map((r) => s.cells[`A${r}`].value)).toEqual(['North', 'south ', 'East', 'South', 'West', 'North'])
  })
})

describe.each(KINDS)('a filter in %s', (kind) => {
  it('hides and shows rows; SUBTOTAL counts the rows shown; undo and redo', async () => {
    const s = await (await opener(kind))(book(), { clock })
    const r = await s.editStructure({ op: 'applyFilter', sheet: 'Data', criteria: south })
    expect(r.ok).toBe(true)
    expect((await s.getCell('Data', 'D1'))?.value).toBe(6)
    expect(hidden((await s.getLayout()).sheets[0])).toEqual([2, 4, 6, 7])
    await s.undo()
    expect((await s.getCell('Data', 'D1'))?.value).toBe(21)
    expect((await s.getLayout()).sheets[0].autoFilter).toEqual({ range: 'A1:B7' })
    await s.redo()
    expect((await s.getCell('Data', 'D1'))?.value).toBe(6)
    await s.editStructure({ op: 'clearFilter', sheet: 'Data' })
    expect((await s.getCell('Data', 'D1'))?.value).toBe(21)
    expect(s.history.undo).toBe(2)
    await s.close()
  })
})
