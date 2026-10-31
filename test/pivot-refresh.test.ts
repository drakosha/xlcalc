/**
 * A session's `refreshPivots`: refused while `xlcalc/pivot` is not loaded; once `registerPivot` loaded it, the cache
 * refreshed from the source's computed values and the PivotTable laid out on its sheet, in a session, the inline
 * calculator and a pool's threads, undone and redone as one step; edits of the source leave it as it was until the next
 * refresh, as Excel does.
 */
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, editStructure, InputError, registerPivot, type BookInput, type CellValue } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
// ru-RU loaded, for a workbook in ru-RU
import './locales.js'
import { closedAfter } from './fixtures/pool.js'

const SOURCE: (string | number | null)[][] = [
  ['Region', 'Product', 'Amount', 'Qty'],
  ['North', 'Apple', 10, 1],
  ['South', 'Banana', 20, 2],
  ['north', 'Cherry', 'n/a', 3],
  ['East', 'Apple', null, 4],
  ['South', 'Apple', 5, 5],
  ['East', 'Banana', 1, 6],
  [null, 'Cherry', 7, 7],
]

const cells = Object.fromEntries(SOURCE.flatMap((row, r) => row.flatMap((v, c) => (v === null ? [] : [[`${'ABCD'[c]}${r + 1}`, { value: v }]]))))
// Qty of the last row by a formula: refresh takes the computed value
cells.D8 = { formula: '3+4' } as never

/** Region on rows, Sum of Qty, the PivotTable not yet laid out at P!A3, over a cache that knows the source's fields. */
const BOOK: BookInput = {
  sheets: [
    { name: 'D', cells },
    { name: 'P', cells: {}, pivotTables: [{ name: 'PT1', cache: 1, location: { range: 'A3:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: [{}, {}, {}, {}], rows: [0], values: [{ field: 3 }] }] },
  ],
  pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:D8' }, fields: ['Region', 'Product', 'Amount', 'Qty'].map((name) => ({ name })) }],
}

/** Excel's compact form: the header, the regions ascending, the blank last, the grand total. */
const LAID: CellValue[][] = [
  ['Row Labels', 'Sum of Qty'],
  ['East', 10],
  ['North', 4],
  ['South', 7],
  ['(blank)', 7],
  ['Grand Total', 28],
]
const EMPTY = LAID.map((r) => r.map(() => null))
const REFRESH = { op: 'refreshPivots' } as const

describe('refreshPivots', () => {
  it('is refused without xlcalc/pivot, and registerPivot takes its engine alone', () => {
    const r = editStructure(BOOK, REFRESH)
    expect(r).toEqual({ ok: false, code: 'pivot', message: 'PivotTables are laid out by xlcalc/pivot, which is not loaded: registerPivot(pivotEngine) loads it' })
    expect(() => registerPivot({ kind: 'xlcalc-pivot' } as never)).toThrow(new InputError('engine', "expected pivotEngine of 'xlcalc/pivot'"))
    registerPivot(pivotEngine)
  })

  for (const execution of ['closures', 'eval'] as const)
    it(`lays the PivotTable out in a session (${execution}), undone and redone; an edit of the source waits for the next refresh`, () => {
      const s = createSession(BOOK, { execution })
      expect(s.editStructure(REFRESH).ok).toBe(true)
      expect(s.getValues('P', 'A3:B8')).toEqual(LAID)
      s.setCells([{ sheet: 'D', cell: 'D2', input: { value: 11 } }])
      expect(s.getValues('P', 'B5')).toEqual([[4]])
      s.editStructure({ op: 'refreshPivots', pivots: [{ sheet: 'p', name: 'pt1' }] })
      expect(s.getValues('P', 'B5:B8')).toEqual([[14], [7], [7], [38]])
      s.undo()
      s.undo()
      expect(s.getValues('P', 'B5')).toEqual([[4]])
      s.undo()
      expect(s.getValues('P', 'A3:B8')).toEqual(EMPTY)
      s.redo()
      expect(s.getValues('P', 'A3:B8')).toEqual(LAID)
      s.close()
    })

  it('refreshes both PivotTables of a cache as one step, and none when one is refused', () => {
    const second = { ...BOOK.sheets[1].pivotTables![0], name: 'PT2', location: { ...BOOK.sheets[1].pivotTables![0].location, range: 'D3:E4' }, rows: [1] }
    const s = createSession({ ...BOOK, sheets: [BOOK.sheets[0], { ...BOOK.sheets[1], pivotTables: [...BOOK.sheets[1].pivotTables!, second] }] })
    expect(s.editStructure({ op: 'refreshPivots', pivots: [{ sheet: 'P', name: 'PT1' }, { sheet: 'P', name: 'PT9' }] }).ok).toBe(false)
    expect(s.getValues('P', 'A4:D4')).toEqual([[null, null, null, null]])
    expect(s.editStructure({ op: 'refreshPivots', pivots: [{ sheet: 'P', name: 'PT2' }] }).ok).toBe(true)
    // PT2 by Product (Apple, Banana, Cherry), and PT1 over the same cache
    expect([s.getValues('P', 'A8:B8'), s.getValues('P', 'D7:E7')]).toEqual([[['Grand Total', 28]], [['Grand Total', 28]]])
    s.undo()
    expect([s.getValues('P', 'A8:B8'), s.getValues('P', 'D7:E7')]).toEqual([[[null, null]], [[null, null]]])
    s.close()
  })

  it('opens a workbook whose cache refreshes on opening as saved: the host refreshes it', () => {
    const book = { ...BOOK, pivotCaches: [{ ...BOOK.pivotCaches![0], refreshOnLoad: true }] }
    const s = createSession(book)
    expect(s.getValues('P', 'A3:B8')).toEqual(EMPTY)
    expect(s.getBook().pivotCaches?.[0].refreshOnLoad).toBe(true)
    s.editStructure(REFRESH)
    expect(s.getValues('P', 'A3:B8')).toEqual(LAID)
    expect(s.getBook().pivotCaches?.[0].refreshOnLoad).toBe(true)
    s.close()
  })

  it("keeps the cells beside its filters, and shows a page field's date item by its serial number", () => {
    // Region across the columns, wider than the Filters area; a date page field (Day) of two dates, the first alone shown
    const cells = { ...BOOK.sheets[0].cells, E1: { value: 'Day' }, ...Object.fromEntries([2, 3, 4, 5, 6, 7, 8].map((r) => [`E${r}`, { value: r < 5 ? 45658 : 45659, numberFormat: 14 }])) }
    const pivot = { ...BOOK.sheets[1].pivotTables![0], location: { range: 'A5:B6', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, pageRows: 2, pageColumns: 1 }, fields: [{}, {}, {}, {}, { multipleItems: true }], rows: undefined, columns: [0], pages: [{ field: 1 }, { field: 4 }] }
    const book: BookInput = {
      sheets: [{ name: 'D', cells }, { name: 'P', cells: { D1: { value: 'kept' }, D3: { value: 'kept' }, E2: { value: 7 } }, pivotTables: [pivot] }],
      pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:E8' }, fields: ['Region', 'Product', 'Amount', 'Qty', 'Day'].map((name) => ({ name })) }],
    }
    const first = editStructure(book, REFRESH)
    if (!first.ok) throw new Error(first.message)
    // the second date hidden, as a page field of several items with one left shown
    const laid = first.book.sheets[1].pivotTables![0]
    const day = { ...laid.fields[4], items: laid.fields[4].items!.map((x, i) => (i === 1 ? { ...x, hidden: true } : x)) }
    const r = editStructure({ ...first.book, sheets: [first.book.sheets[0], { ...first.book.sheets[1], pivotTables: [{ ...laid, fields: laid.fields.map((f, i) => (i === 4 ? day : f)) }] }] }, REFRESH)
    if (!r.ok) throw new Error(r.message)
    const p = r.book.sheets[1].cells
    expect([p.A3?.value, p.B3?.value, p.B3?.numberFormat, p.D3, p.E2]).toEqual(['Day', 45658, 14, { value: 'kept' }, { value: 7 }])
  })

  it("sorts the month names of the workbook's locale by its list", () => {
    const months = ['март', 'январь', 'февраль', 'апрель']
    const cells = { A1: { value: 'Month' }, B1: { value: 'N' }, ...Object.fromEntries(months.flatMap((m, i) => [[`A${i + 2}`, { value: m }], [`B${i + 2}`, { value: i + 1 }]])) }
    const book: BookInput = {
      locale: 'ru-RU',
      sheets: [{ name: 'D', cells }, { name: 'P', cells: {}, pivotTables: [{ name: 'PT1', cache: 1, location: { range: 'A3:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: [{ sort: { order: 'ascending' } }, {}], rows: [0], values: [{ field: 1 }] }] }],
      pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:B5' }, fields: [{ name: 'Month' }, { name: 'N' }] }],
    }
    const r = editStructure(book, REFRESH)
    if (!r.ok) throw new Error(r.message)
    expect([4, 5, 6, 7].map((n) => r.book.sheets[1].cells[`A${n}`]?.value)).toEqual(['январь', 'февраль', 'март', 'апрель'])
    // in an en-US workbook they are text, in Excel's order of text
    const en = editStructure({ ...book, locale: undefined }, REFRESH)
    if (!en.ok) throw new Error(en.message)
    expect([4, 5, 6, 7].map((n) => en.book.sheets[1].cells[`A${n}`]?.value)).toEqual(['апрель', 'март', 'февраль', 'январь'])
  })

  it('refuses a PivotTable not there, and leaves the book', () => {
    const s = createSession(BOOK)
    const r = s.editStructure({ op: 'refreshPivots', pivots: [{ sheet: 'P', name: 'PT9' }] })
    expect(r).toMatchObject({ ok: false, code: 'pivot', message: 'there is no PivotTable PT9 on P' })
    expect(s.getValues('P', 'A3:B4')).toEqual([[null, null], [null, null]])
    s.close()
  })

  it('reads a source by a table, a name, and a name whose formula gives the reference', () => {
    const sources: [BookInput['pivotCaches'], Partial<BookInput>][] = [
      [[{ ...BOOK.pivotCaches![0], source: { type: 'table', table: 'sales' } }], { sheets: [{ ...BOOK.sheets[0], tables: [{ name: 'Sales', range: 'A1:D8', columns: SOURCE[0].map((name) => ({ name: String(name) })) }] }, BOOK.sheets[1]] }],
      [[{ ...BOOK.pivotCaches![0], source: { type: 'name', name: 'Data' } }], { names: [{ name: 'Data', ref: 'D!$A$1:$D$8' }] }],
      [[{ ...BOOK.pivotCaches![0], source: { type: 'name', name: 'Grown' } }], { names: [{ name: 'Grown', ref: 'OFFSET(D!$A$1,0,0,COUNTA(D!$B:$B),4)' }] }],
    ]
    for (const [caches, more] of sources) {
      const r = editStructure({ ...BOOK, ...more, pivotCaches: caches }, { op: 'refreshPivots', pivots: [{ sheet: 'P', name: 'PT1' }] })
      expect([caches![0].source, r.ok && r.book.sheets[1].cells.B8]).toEqual([caches![0].source, { value: 28 }])
    }
  })

  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`lays it out in the ${kind} calculator, undone`, async () => {
      const c = await calc()
      const s = await c.createSession(BOOK)
      expect((await s.editStructure(REFRESH)).ok).toBe(true)
      expect(await s.getValues('P', 'A3:B8')).toEqual(LAID)
      await s.undo()
      expect(await s.getValues('P', 'A3:B8')).toEqual(EMPTY)
      await s.close()
    })

  // registerPivot takes an entry into a PivotTable's cell too: an item's label typed anew is an edit of its definition (editPivotTable)
  for (const [kind, calc] of [['a session', async () => null], ['the inline calculator', async () => createInlineCalc()], ['a pool', () => made({ threads: 1 })]] as const)
    it(`takes an entry into an item's label as editPivotTable in ${kind}`, async () => {
      const c = await calc()
      const s = c ? await c.createSession(BOOK) : createSession(BOOK)
      await s.editStructure(REFRESH)
      expect(await s.setCellText('P', 'A4', 'Восток')).toMatchObject({ ok: true, changes: [{ sheet: 'P', cell: 'A4', value: 'Восток' }] })
      expect(await s.setCellText('P', 'A4', 'восток')).toMatchObject({ ok: true, changes: [] })
      expect(await s.setCellText('P', 'B4', '3')).toMatchObject({ ok: false, code: 'pivot' })
      await s.close()
    })
})
