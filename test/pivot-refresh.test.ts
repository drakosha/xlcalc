/**
 * A session's `refreshPivots`: refused while `xlcalc/pivot` is not loaded; once `registerPivot` loaded it, the cache
 * refreshed from the source's computed values and the PivotTable laid out on its sheet, in a session, the inline
 * calculator and a pool's threads, undone and redone as one step; edits of the source leave it as it was until the next
 * refresh, as Excel does.
 */
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, editStructure, InputError, registerPivot, type BookInput, type CellValue } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
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

  it('lays the PivotTable out in a session, undone and redone; an edit of the source waits for the next refresh', () => {
    const s = createSession(BOOK)
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

  it('refuses a PivotTable not there, and leaves the book', () => {
    const s = createSession(BOOK)
    const r = s.editStructure({ op: 'refreshPivots', pivots: [{ sheet: 'P', name: 'PT9' }] })
    expect(r).toMatchObject({ ok: false, code: 'pivot', message: 'there is no PivotTable PT9 on P' })
    expect(s.getValues('P', 'A3:B4')).toEqual([[null, null], [null, null]])
    s.close()
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
})
