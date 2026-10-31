/**
 * The operations on a pivot cache (`editStructure`): dates grouped by periods, numbers by an interval, items by hand,
 * ungrouped, a group renamed; calculated fields and items added, changed, deleted. Each acts on both PivotTables of the
 * cache and lays them out in one step, undone and redone, in a session (both modes), the inline calculator and a pool.
 */
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, editStructure, registerPivot, type BookInput, type CellValue, type StructureOp } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { closedAfter } from './fixtures/pool.js'

const ROWS: [string, string, number, number][] = [
  ['North', '2024-01-15', 1, 10],
  ['South', '2024-02-10', 2, 20],
  ['East', '2024-02-20', 3, 30],
  ['North', '2025-01-05', 4, 40],
  ['West', '2025-03-01', 5, 50],
]
const cells: BookInput['sheets'][number]['cells'] = { A1: { value: 'Region' }, B1: { value: 'Date' }, C1: { value: 'Qty' }, D1: { value: 'Price' } }
ROWS.forEach(([region, date, qty, price], i) => {
  cells[`A${i + 2}`] = { value: region }
  cells[`B${i + 2}`] = { value: { date }, numberFormat: 'yyyy-mm-dd' }
  cells[`C${i + 2}`] = { value: qty }
  cells[`D${i + 2}`] = { value: price }
})

const table = (name: string, at: string, rows: number[], value: number) => ({ name, cache: 1, location: { range: at, firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: [{}, {}, {}, {}], rows, values: [{ field: value }] })

/** PT1 (Sum of Qty) at P!A3 and PT2 (Sum of Price) at P!E3, both by the field `rows`, over one cache, laid out. */
function book(rows: number): BookInput {
  registerPivot(pivotEngine)
  const b: BookInput = {
    sheets: [
      { name: 'D', cells },
      { name: 'P', cells: {}, pivotTables: [table('PT1', 'A3', [rows], 2), table('PT2', 'E3', [rows], 3)] },
    ],
    pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:D6' }, fields: ['Region', 'Date', 'Qty', 'Price'].map((name) => ({ name })) }],
  }
  const r = editStructure(b, { op: 'refreshPivots' })
  if (!r.ok) throw new Error(r.message)
  return r.book
}

const P = { sheet: 'P', pivot: 'PT1' } as const

const byMonths: StructureOp = { op: 'groupPivotField', ...P, field: 1, group: { by: 'dates', periods: ['months', 'years'] } }
/** PT1 by years and months; PT2 by years, months and the dates. */
const PT1_MONTHS: CellValue[][] = [['Row Labels', 'Sum of Qty'], ['2024', 6], ['Jan', 1], ['Feb', 5], ['2025', 9], ['Jan', 4], ['Mar', 5], ['Grand Total', 15]]

describe('grouping a field of a PivotTable', () => {
  for (const execution of ['closures', 'eval'] as const)
    it(`groups dates by months and years on both PivotTables of the cache, undone and redone (${execution})`, () => {
      const s = createSession(book(1), { execution })
      const before = s.getValues('P', 'A3:B9')
      expect(s.editStructure(byMonths).ok).toBe(true)
      expect(s.getValues('P', 'A3:B10')).toEqual(PT1_MONTHS)
      const fields = s.getBook().pivotCaches![0].fields
      expect(fields.map((f) => f.name)).toEqual(['Region', 'Date', 'Qty', 'Price', 'Months (Date)', 'Years (Date)'])
      expect(fields[4].group?.items).toEqual(['<1/15/2024', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', '>3/2/2025'])
      // the other PivotTable shows the periods before the dates
      const pt2 = s.getBook().sheets[1].pivotTables!.find((t) => t.name === 'PT2')!
      expect(pt2.rows).toEqual([5, 4, 1])
      expect(s.getValues('P', 'E4:F5')).toEqual([['2024', 60], ['Jan', 10]])
      s.undo()
      expect(s.getValues('P', 'A3:B9')).toEqual(before)
      s.redo()
      expect(s.getValues('P', 'A3:B10')).toEqual(PT1_MONTHS)
      // ungrouped from the other PivotTable: the periods gone from both, the dates back where they were
      expect(s.editStructure({ op: 'ungroupPivotField', sheet: 'P', pivot: 'PT2', field: 4 }).ok).toBe(true)
      expect(s.getBook().pivotCaches![0].fields.length).toBe(4)
      expect(s.getBook().sheets[1].pivotTables!.map((t) => t.rows)).toEqual([[1], [1]])
      expect(s.getValues('P', 'A3:B9')).toEqual(before)
      s.close()
    })

  it('groups days by a number of days from a start, and numbers by an interval', () => {
    const s = createSession(book(1))
    s.editStructure({ op: 'groupPivotField', ...P, field: 1, group: { by: 'dates', periods: ['days'], days: 30, start: '2024-01-01' } })
    expect(s.getValues('P', 'A4:B7')).toEqual([['1/1/2024 - 1/30/2024', 1], ['1/31/2024 - 2/29/2024', 5], ['12/26/2024 - 1/24/2025', 4], ['2/24/2025 - 3/2/2025', 5]])
    const n = createSession(book(2))
    expect(n.editStructure({ op: 'groupPivotField', ...P, field: 2, group: { by: 'numbers', interval: 2 } }).ok).toBe(true)
    expect(n.getValues('P', 'A3:B7')).toEqual([['Row Labels', 'Sum of Qty'], ['1-2', 3], ['3-5', 12], ['Grand Total', 15], [null, null]])
    s.close()
    n.close()
  })

  it('groups items by hand, renames a group in the cache, and ungroups it', () => {
    const s = createSession(book(0))
    // Region on PT1 lists East, North, South, West: East and West into Group1
    expect(s.editStructure({ op: 'groupPivotField', ...P, field: 0, group: { by: 'items', items: [0, 3] } }).ok).toBe(true)
    expect(s.getValues('P', 'A3:B10')).toEqual([['Row Labels', 'Sum of Qty'], ['Group1', 8], ['East', 3], ['West', 5], ['North', 5], ['North', 5], ['South', 2], ['South', 2]])
    expect(s.getBook().sheets[1].pivotTables!.map((t) => t.rows)).toEqual([[4, 0], [4, 0]])
    expect(s.editStructure({ op: 'renamePivotGroup', ...P, field: 4, item: 0, name: 'Coast' }).ok).toBe(true)
    expect([s.getValues('P', 'A4'), s.getValues('P', 'E4')]).toEqual([[['Coast']], [['Coast']]])
    expect(s.editStructure({ op: 'renamePivotGroup', ...P, field: 4, item: 0, name: 'north' })).toMatchObject({ ok: false, code: 'pivot', message: 'PivotTable item name already exists: "north"' })
    expect(s.editStructure({ op: 'ungroupPivotField', ...P, field: 4 }).ok).toBe(true)
    expect(s.getValues('P', 'A3:B8')).toEqual([['Row Labels', 'Sum of Qty'], ['East', 3], ['North', 5], ['South', 2], ['West', 5], ['Grand Total', 15]])
    s.close()
  })

  it('refuses text among the dates, a field not grouped, and checks its input', () => {
    const b = book(0)
    expect(editStructure(b, { op: 'groupPivotField', ...P, field: 0, group: { by: 'dates', periods: ['months'] } })).toMatchObject({ ok: false, code: 'pivot' })
    expect(editStructure(b, { op: 'ungroupPivotField', ...P, field: 0 })).toMatchObject({ ok: false, message: 'The field "Region" is not grouped' })
    expect(() => editStructure(b, { op: 'groupPivotField', ...P, field: 9, group: { by: 'numbers', interval: 1 } })).toThrow('op.field')
    expect(() => editStructure(b, { op: 'groupPivotField', sheet: 'P', pivot: 'PT9', field: 1, group: { by: 'numbers', interval: 1 } })).toThrow('unknown PivotTable "PT9" on "P"')
    expect(() => editStructure(b, { op: 'groupPivotField', ...P, field: 1, group: { by: 'dates', periods: ['weeks'] } } as never)).toThrow('op.group.periods[0]')
    expect(() => editStructure(b, { op: 'groupPivotField', ...P, field: 2, group: { by: 'numbers', interval: 0 } })).toThrow('op.group.interval')
  })
})

describe('grouping with slicers', () => {
  it('filters PivotTables grouped by months with a slicer of another field, and with a slicer of the years', () => {
    const s = createSession(book(1))
    s.editStructure(byMonths)
    const pivots = [{ sheet: 'P', name: 'PT1' }]
    expect(s.editStructure({ op: 'addSlicer', sheet: 'P', slicer: {}, cache: { field: 'Region', pivots } }).ok).toBe(true)
    // Region's items in the cache: North, South, East, West
    s.editStructure({ op: 'editSlicer', slicer: 'Region', cache: { items: [0, 1, 2, 3].map((item) => ({ item, selected: item === 0 })) } })
    expect(s.getValues('P', 'A3:B8')).toEqual([['Row Labels', 'Sum of Qty'], ['2024', 1], ['Jan', 1], ['2025', 4], ['Jan', 4], ['Grand Total', 5]])
    s.editStructure({ op: 'editSlicer', slicer: 'Region', cache: { items: [0, 1, 2, 3].map((item) => ({ item })) } })
    // the years' groups: <1/15/2024, 2024, 2025, >3/2/2025
    expect(s.editStructure({ op: 'addSlicer', sheet: 'P', slicer: {}, cache: { field: 'Years (Date)', pivots } }).ok).toBe(true)
    s.editStructure({ op: 'editSlicer', slicer: 'Years (Date)', cache: { items: [0, 1, 2, 3].map((item) => ({ item, selected: item === 2 })) } })
    expect(s.getValues('P', 'A3:B7')).toEqual([['Row Labels', 'Sum of Qty'], ['2025', 9], ['Jan', 4], ['Mar', 5], ['Grand Total', 9]])
    s.close()
  })
})

describe('calculated fields and items', () => {
  for (const execution of ['closures', 'eval'] as const)
    it(`adds a calculated field over the sums, changes and deletes it, on both PivotTables (${execution})`, () => {
      const s = createSession(book(0), { execution })
      expect(s.editStructure({ op: 'setPivotCalculatedField', ...P, name: 'Unit price', formula: '=Price/Qty' }).ok).toBe(true)
      const pt = (b: BookInput) => b.sheets[1].pivotTables!
      // not placed on Values by the operation; placed, every cell is the formula over the sums: Price/Qty is 10 throughout
      const b = s.getBook()
      expect(b.sheets[1].pivotTables![0].values).toEqual([{ field: 2 }])
      const placed = editStructure({ ...b, sheets: [b.sheets[0], { ...b.sheets[1], pivotTables: b.sheets[1].pivotTables!.map((t) => ({ ...t, values: [{ field: 4, name: 'Sum of Unit price' }] })) }] }, { op: 'refreshPivots' })
      expect(placed.ok && [4, 5, 6, 7, 8].map((r) => placed.book.sheets[1].cells[`B${r}`]?.value)).toEqual([10, 10, 10, 10, 10])
      expect(s.editStructure({ op: 'setPivotCalculatedField', ...P, name: 'unit PRICE', formula: "=Price/Qty*2+'Qty'" }).ok).toBe(true)
      expect(s.getBook().pivotCaches![0].fields[4]).toEqual({ name: 'Unit price', formula: "Price/Qty*2+'Qty'" })
      expect(s.editStructure({ op: 'setPivotCalculatedField', ...P, name: 'Unit price', formula: null }).ok).toBe(true)
      expect(s.getBook().pivotCaches![0].fields.length).toBe(4)
      s.undo()
      expect(s.getBook().pivotCaches![0].fields.length).toBe(5)
      s.close()
    })

  it('adds a calculated item counted in the grand total, in the file form, and deletes it', () => {
    const s = createSession(book(0))
    expect(s.editStructure({ op: 'setPivotCalculatedItem', ...P, field: 0, name: 'Coast', formula: "=East+'West'" }).ok).toBe(true)
    expect(s.getBook().pivotCaches![0].calculatedItems).toEqual([{ field: 0, item: 4, formula: 'Region[East]+Region[West]' }])
    expect(s.getValues('P', 'A3:B9')).toEqual([['Row Labels', 'Sum of Qty'], ['East', 3], ['North', 5], ['South', 2], ['West', 5], ['Coast', 8], ['Grand Total', 23]])
    expect(s.getValues('P', 'E8:F9')).toEqual([['Coast', 80], ['Grand Total', 230]])
    expect(s.editStructure({ op: 'setPivotCalculatedItem', ...P, field: 0, name: 'coast', formula: null }).ok).toBe(true)
    expect(s.getValues('P', 'A8:B8')).toEqual([['Grand Total', 15]])
    expect(s.editStructure({ op: 'groupPivotField', ...P, field: 0, group: { by: 'items', items: [0] } }).ok).toBe(true)
    expect(s.editStructure({ op: 'setPivotCalculatedItem', ...P, field: 0, name: 'Coast', formula: '=East' })).toMatchObject({ ok: false, code: 'pivot' })
    s.close()
  })

  it('refuses a reference to a cell and a formula that does not read', () => {
    const b = book(0)
    expect(editStructure(b, { op: 'setPivotCalculatedField', ...P, name: 'X', formula: '=A1+Qty' })).toMatchObject({ ok: false, message: 'References, names, and arrays are not supported in PivotTable formulas' })
    expect(editStructure(b, { op: 'setPivotCalculatedField', ...P, name: 'X', formula: '=Qty*(' })).toMatchObject({ ok: false, message: 'The formula "=Qty*(" contains an error' })
    expect(editStructure(b, { op: 'setPivotCalculatedField', ...P, name: 'X', formula: '=X+1' })).toMatchObject({ ok: false, code: 'pivot' })
    expect(() => editStructure(b, { op: 'setPivotCalculatedField', ...P, name: 'Y', formula: null })).toThrow('unknown calculated field "Y"')
  })
})

describe('the operations in the other kinds of session', () => {
  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`groups and adds a calculated item in the ${kind} calculator, undone`, async () => {
      const c = await calc()
      const s = await c.createSession(book(1))
      expect((await s.editStructure(byMonths)).ok).toBe(true)
      expect(await s.getValues('P', 'A3:B10')).toEqual(PT1_MONTHS)
      await s.undo()
      expect(await s.getValues('P', 'A4')).not.toEqual([['2024']])
      await s.close()
    })
})
