/**
 * Slicers and timelines through `editStructure`: added over a new cache of a field or a cache of the workbook, their
 * choice the filter of every PivotTable they filter (hidden items, a date filter), changed, connected to another
 * PivotTable, deleted with their cache; refused and checked as input. Without `xlcalc/pivot` the PivotTables keep their
 * cells and their cache is marked stale; with it they are laid out again from the cache, in a session, the inline
 * calculator and a pool's threads, both execution modes, undone and redone as one step.
 */
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, editStructure, InputError, registerPivot, type BookInput, type CellValue, type StructureOp } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { closedAfter } from './fixtures/pool.js'
import { pivotBook } from './pivot-books.js'

const day = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86_400_000 + 25_569
const SOURCE: [string | null, number, string, number][] = [
  ['North', day(2024, 1, 5), 'Apple', 10],
  ['South', day(2024, 6, 1), 'Banana', 20],
  ['East', day(2025, 2, 1), 'Apple', 5],
  ['North', day(2025, 3, 9), 'Cherry', 7],
  ['West', day(2025, 11, 30), 'Banana', 1],
]
const cells: BookInput['sheets'][number]['cells'] = { A1: { value: 'Region' }, B1: { value: 'Date' }, C1: { value: 'Product' }, D1: { value: 'Amount' } }
SOURCE.forEach(([region, date, product, amount], i) => {
  if (region) cells[`A${i + 2}`] = { value: region }
  cells[`B${i + 2}`] = { value: date, numberFormat: 'yyyy-mm-dd' }
  cells[`C${i + 2}`] = { value: product }
  cells[`D${i + 2}`] = { value: amount }
})

const location = (range: string) => ({ range, firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 })
/** Sum of Amount by Region at P!A3, by Product at P!E3, over one cache; refreshed before each test's operations. */
const BOOK: BookInput = {
  sheets: [
    { name: 'D', cells },
    {
      name: 'P',
      cells: {},
      pivotTables: [
        { name: 'PT1', cache: 1, location: location('A3:B4'), fields: [{}, {}, {}, {}], rows: [0], values: [{ field: 3 }] },
        { name: 'PT2', cache: 1, location: location('E3:F4'), fields: [{}, {}, {}, {}], rows: [2], values: [{ field: 3 }] },
      ],
    },
  ],
  pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:D6' }, fields: ['Region', 'Date', 'Product', 'Amount'].map((name) => ({ name })) }],
}

const PT1 = { sheet: 'P', name: 'PT1' }
const PT2 = { sheet: 'P', name: 'PT2' }
const REGION: StructureOp = { op: 'addSlicer', sheet: 'P', slicer: {}, cache: { field: 'Region', pivots: [PT1] } }
/** Region's items in the cache, as the refresh lists them: North, South, East, West. */
const choose = (...items: number[]): StructureOp => ({ op: 'editSlicer', slicer: 'Region', cache: { items: [0, 1, 2, 3].map((item) => (items.includes(item) ? { item } : { item, selected: false })) } })

const ALL: CellValue[][] = [['Row Labels', 'Sum of Amount'], ['East', 5], ['North', 17], ['South', 20], ['West', 1], ['Grand Total', 43]]
const NORTH: CellValue[][] = [['Row Labels', 'Sum of Amount'], ['North', 17], ['Grand Total', 17], [null, null], [null, null], [null, null]]

const fresh = () => {
  const r = editStructure(BOOK, { op: 'refreshPivots' })
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe('slicers and timelines without xlcalc/pivot', () => {
  it('are refused a refresh, and write their choice with the cache marked stale, the cells kept', () => {
    expect(editStructure(BOOK, { op: 'refreshPivots' }).ok).toBe(false)
    // Sum of Amount by Region at Pivot!A3, its cache with its records
    const book = pivotBook()
    delete book.slicerCaches
    delete book.sheets[1].slicers
    const at = { sheet: 'Pivot', name: 'PivotTable1' }
    const r = editStructure(book, { op: 'addSlicer', sheet: 'Pivot', slicer: {}, cache: { field: 'Amount', pivots: [at] } })
    if (!r.ok) throw new Error(r.message)
    // the field listed in the cache as Excel lists a slicer's, nothing laid out
    expect([r.book.pivotCaches![0].fields[1].items, r.book.pivotCaches![0].records![1], r.stale]).toEqual([[10, 20, 30, 40, 50, 60, 70, 80], [1, 1, 0], undefined])
    const c = editStructure(r.book, { op: 'editSlicer', slicer: 'Amount', cache: { items: [{ item: 0, selected: false }] } })
    if (!c.ok) throw new Error(c.message)
    expect(c.book.sheets[1].pivotTables![0].fields[1].items).toEqual([{ item: 0, hidden: true }, ...[1, 2, 3, 4, 5, 6, 7].map((item) => ({ item }))])
    expect([c.book.pivotCaches![0].stale, c.stale, c.book.sheets[1].cells.B8]).toEqual([true, [at], { value: 360 }])
  })
})

describe('slicers and timelines with xlcalc/pivot', () => {
  it('loads', () => registerPivot(pivotEngine))

  for (const execution of ['closures', 'eval'] as const)
    it(`a slicer filters its PivotTables in a session (${execution}), undone and redone one step each`, () => {
      const s = createSession(fresh(), { execution })
      expect(s.getValues('P', 'A3:B8')).toEqual(ALL)
      expect(s.editStructure(REGION).ok).toBe(true)
      const b = s.getBook()
      expect(b.sheets[1].slicers).toEqual([{ caption: 'Region', name: 'Region', cache: 'Slicer_Region', place: { from: 'D3', to: 'G17' } }])
      // East, North, South, West: in the slicer's order, as Excel saves them
      expect(b.slicerCaches![0].items).toEqual([{ item: 2 }, { item: 0 }, { item: 1 }, { item: 3 }])
      expect(s.editStructure(choose(0)).ok).toBe(true)
      expect(s.getValues('P', 'A3:B8')).toEqual(NORTH)
      // the other regions under the choice of North have their data still: Region is the slicer's own field
      expect(s.getBook().slicerCaches![0].items).toEqual([{ item: 2, selected: false }, { item: 0 }, { item: 1, selected: false }, { item: 3, selected: false }])
      s.undo()
      expect(s.getValues('P', 'A3:B8')).toEqual(ALL)
      s.redo()
      expect(s.getValues('P', 'A3:B8')).toEqual(NORTH)
      s.undo()
      s.undo()
      expect(s.getBook().slicerCaches).toBeUndefined()
      s.close()
    })

  it("a slicer of a field on no axis hides its items there; another slicer's items without data go last, marked", () => {
    const s = createSession(fresh())
    s.editStructure({ op: 'addSlicer', sheet: 'P', slicer: { name: 'Fruit' }, cache: { field: 'product', pivots: [PT1] } })
    s.editStructure(REGION)
    s.editStructure({ op: 'editSlicer', slicer: 'Fruit', cache: { items: [{ item: 0 }, { item: 1, selected: false }, { item: 2, selected: false }] } })
    // Apple: North 10, East 5
    expect(s.getValues('P', 'A3:B6')).toEqual([['Row Labels', 'Sum of Amount'], ['East', 5], ['North', 10], ['Grand Total', 15]])
    const b = s.getBook()
    expect(b.sheets[1].pivotTables![0].fields[2].items).toEqual([{ item: 0 }, { item: 1, hidden: true }, { item: 2, hidden: true }])
    expect(b.slicerCaches!.find((c) => c.name === 'Slicer_Region')!.items).toEqual([{ item: 2 }, { item: 0 }, { item: 1, noData: true }, { item: 3, noData: true }])
    s.close()
  })

  it('takes the filter a PivotTable has, filters a PivotTable connected later, and keeps the filter when deleted', () => {
    const s = createSession(fresh())
    s.editStructure({ op: 'addSlicer', sheet: 'P', slicer: { name: 'Fruit' }, cache: { field: 'Product', pivots: [PT2] } })
    s.editStructure({ op: 'editSlicer', slicer: 'Fruit', cache: { items: [{ item: 0, selected: false }, { item: 1 }, { item: 2 }] } })
    // a second slicer over the field takes the choice its PivotTable holds
    s.editStructure({ op: 'addSlicer', sheet: 'P', slicer: { name: 'Fruit 2' }, cache: { field: 'Product', pivots: [PT2] } })
    expect(s.getBook().slicerCaches![1].items!.find((x) => x.item === 0)).toEqual({ item: 0, selected: false })
    s.editStructure({ op: 'deleteSlicer', slicer: 'Fruit 2' })
    expect(s.editStructure({ op: 'editSlicer', slicer: 'Fruit', cache: { pivots: [PT2, PT1] } }).ok).toBe(true)
    // PT1 without Apple: North 7, South 20, West 1
    expect(s.getValues('P', 'A4:B7')).toEqual([['North', 7], ['South', 20], ['West', 1], ['Grand Total', 28]])
    expect(s.editStructure({ op: 'deleteSlicer', slicer: 'Fruit' }).ok).toBe(true)
    const b = s.getBook()
    expect([b.slicerCaches, b.sheets[1].slicers, b.sheets[1].pivotTables![0].fields[2].items?.[0]]).toEqual([undefined, undefined, { item: 0, hidden: true }])
    s.close()
  })

  it('a timeline filters its PivotTables by its dates, of whole days, and takes the level they span', () => {
    const s = createSession(fresh())
    expect(s.editStructure({ op: 'addTimeline', sheet: 'P', timeline: {}, cache: { field: 'Date', pivots: [PT1, PT2] } }).ok).toBe(true)
    expect(s.getBook().timelineCaches).toEqual([{ name: 'NativeTimeline_Date', field: 'Date', pivots: [PT1, PT2], cache: 1, bounds: { start: '2024-01-01T00:00:00', end: '2026-01-01T00:00:00' } }])
    expect(s.editStructure({ op: 'editTimeline', timeline: 'Date', cache: { selection: { start: '2025-01-01T00:00:00', end: '2025-03-31T00:00:00' } } }).ok).toBe(true)
    expect(s.getValues('P', 'A3:B6')).toEqual([['Row Labels', 'Sum of Amount'], ['East', 5], ['North', 7], ['Grand Total', 12]])
    const b = s.getBook()
    expect([b.sheets[1].timelines![0].selectionLevel, b.sheets[1].pivotTables![1].filters]).toEqual(['quarters', [{ field: 1, type: 'dateBetween', args: ['2025-01-01', '2025-03-31'], name: 'Date' }]])
    s.editStructure({ op: 'editTimeline', timeline: 'Date', cache: { selection: { start: '2024-06-01T00:00:00', end: '2024-06-01T00:00:00' } } })
    expect(s.getValues('P', 'A3:B5')).toEqual([['Row Labels', 'Sum of Amount'], ['South', 20], ['Grand Total', 20]])
    expect(s.getBook().sheets[1].timelines![0].selectionLevel).toBe('days')
    // Clear Filter: every date again
    s.editStructure({ op: 'editTimeline', timeline: 'Date', cache: { selection: null } })
    expect(s.getValues('P', 'A3:B8')).toEqual(ALL)
    // deleted, its dates no longer filter, as in Excel (where a slicer's choice stays)
    s.editStructure({ op: 'editTimeline', timeline: 'Date', cache: { selection: { start: '2025-01-01T00:00:00', end: '2025-03-31T00:00:00' } } })
    s.editStructure({ op: 'deleteTimeline', timeline: 'Date' })
    expect([s.getBook().timelineCaches, s.getBook().sheets[1].pivotTables![1].filters]).toEqual([undefined, undefined])
    expect(s.getValues('P', 'A3:B8')).toEqual(ALL)
    s.close()
  })

  it('a refresh keeps the choice: a new item comes unchosen, as in a manual filter', () => {
    const s = createSession(fresh())
    s.editStructure(REGION)
    s.editStructure(choose(0, 1))
    s.setCells([{ sheet: 'D', cell: 'A6', input: { value: 'Central' } }])
    s.editStructure({ op: 'refreshPivots' })
    expect(s.getValues('P', 'A3:B6')).toEqual([['Row Labels', 'Sum of Amount'], ['North', 17], ['South', 20], ['Grand Total', 37]])
    const b = s.getBook()
    const items = b.pivotCaches![0].fields[0].items!
    expect(b.slicerCaches![0].items!.map((x) => [items[x.item], x.selected !== false])).toEqual([['Central', false], ['East', false], ['North', true], ['South', true], ['West', false]])
    s.close()
  })

  it('refuses a field not of the cache, a name taken, a timeline without dates; checks its input', () => {
    const b = fresh()
    const refusal = (op: StructureOp) => {
      const r = editStructure(b, op)
      return r.ok ? null : [r.code, r.message]
    }
    expect(refusal({ ...REGION, cache: { field: 'Month', pivots: [PT1] } } as StructureOp)).toEqual(['pivot', `there is no field "Month" in the PivotTables' source`])
    expect(refusal({ op: 'addTimeline', sheet: 'P', timeline: {}, cache: { field: 'Region', pivots: [PT1] } })).toEqual(['pivot', 'a timeline needs a field of dates, and "Region" holds none'])
    const r = editStructure(b, REGION)
    if (!r.ok) throw new Error(r.message)
    expect(editStructure(r.book, { op: 'addSlicer', sheet: 'P', slicer: { name: 'region' }, cache: 'Slicer_Region' })).toMatchObject({ ok: false, code: 'pivot', message: 'the name "region" is already used by a slicer or a timeline' })
    const throws = (op: unknown, m: string) => expect(() => editStructure(r.book, op as StructureOp)).toThrow(new InputError(...(m.split(': ') as [string, string])))
    throws(choose(), 'op.cache.items: expected at least one item chosen')
    throws({ op: 'editSlicer', slicer: 'Region', cache: { items: [{ item: 9 }] } }, 'op.cache.items[0].item: expected an integer from 0 to 3, got 9')
    throws({ op: 'editSlicer', slicer: 'Nope' }, 'op.slicer: unknown slicer "Nope"')
    throws({ op: 'editSlicer', slicer: 'Region', set: { name: 'X' } }, "op.set: a slicer's name and cache are not edited")
    throws({ op: 'addSlicer', sheet: 'P', slicer: {}, cache: 'Slicer_Nope' }, 'op.cache: unknown slicer cache "Slicer_Nope"')
    throws({ op: 'editTimeline', timeline: 'Region' }, 'op.timeline: unknown timeline "Region"')
    // the new slicer named after its field again: Region 1
    const again = editStructure(r.book, { op: 'addSlicer', sheet: 'P', slicer: {}, cache: 'Slicer_Region' })
    expect(again.ok && again.book.sheets[1].slicers!.map((x) => x.name)).toEqual(['Region', 'Region 1'])
    const two = editStructure(again.ok ? again.book : b, { op: 'deleteSlicer', slicer: 'Region' })
    expect(two.ok && two.book.slicerCaches!.length).toBe(1)
  })

  it("a table's slicer is added over its cache, changed in its look and order, and deleted; its choice stays the table's AutoFilter", () => {
    const book: BookInput = { ...fresh(), slicerCaches: [{ name: 'Slicer_Region1', field: 'Region', pivots: [], table: { name: 'Sales', column: 'Region' } }] }
    book.sheets[0] = { ...book.sheets[0], tables: [{ name: 'Sales', range: 'A1:D6', columns: ['Region', 'Date', 'Product', 'Amount'].map((name) => ({ name })) }], slicers: [{ name: 'Region 2', cache: 'Slicer_Region1' }] }
    const r = editStructure(book, { op: 'addSlicer', sheet: 'D', slicer: { name: 'Again' }, cache: 'Slicer_Region1' })
    if (!r.ok) throw new Error(r.message)
    expect(r.book.sheets[0].slicers).toEqual([{ name: 'Region 2', cache: 'Slicer_Region1' }, { caption: 'Region', name: 'Again', cache: 'Slicer_Region1' }])
    const e = editStructure(r.book, { op: 'editSlicer', slicer: 'Again', set: { style: 'SlicerStyleDark2' }, cache: { sort: 'descending' } })
    expect(e.ok && [e.book.sheets[0].slicers![1].style, e.book.slicerCaches![0].sort]).toEqual(['SlicerStyleDark2', 'descending'])
    expect(() => editStructure(r.book, { op: 'editSlicer', slicer: 'Again', cache: { items: [{ item: 0 }] } })).toThrow("a table's slicer chooses by the table's AutoFilter")
    const d = editStructure(r.book, { op: 'deleteSlicer', slicer: 'Again' })
    const gone = d.ok && editStructure(d.book, { op: 'deleteSlicer', slicer: 'Region 2' })
    expect([d.ok && d.book.slicerCaches!.length, gone && gone.ok && gone.book.slicerCaches]).toEqual([1, undefined])
  })

  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`filters in the ${kind} calculator, undone`, async () => {
      const c = await calc()
      const s = await c.createSession(fresh())
      expect((await s.editStructure(REGION)).ok).toBe(true)
      expect((await s.editStructure(choose(0))).ok).toBe(true)
      expect(await s.getValues('P', 'A3:B8')).toEqual(NORTH)
      await s.undo()
      expect(await s.getValues('P', 'A3:B8')).toEqual(ALL)
      await s.close()
    })
})
