/**
 * GETPIVOTDATA over a workbook's PivotTable (`test/pivot-books.ts`, Sum of Amount by Region): the value field and the
 * pairs as Excel matches them, the grand total, what the layout does not show, a stale cache; the PivotTable's cells as
 * the formula's references in every kind of session and both execution modes, a refresh changing them, with undo; its layout read by the formula
 * at each computation, so a PivotTable laid out otherwise over the same cells answers otherwise. Excel's answers are the
 * set `gpd-1` (`pivot-gpd.test.ts`).
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, registerPivot, type BookInput, type Execution } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { pivotBook } from './pivot-books.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

afterAll(closeCalcs)
// a refresh lays the PivotTable out anew through xlcalc/pivot
registerPivot(pivotEngine)

/** The pivot book with formulas on a sheet G, from A1 down. */
function withFormulas(formulas: string[], change?: (b: BookInput) => void): BookInput {
  const b = pivotBook()
  change?.(b)
  b.sheets.push({ name: 'G', cells: Object.fromEntries(formulas.map((f, i) => [`A${i + 1}`, { formula: f }])) })
  return b
}

const values = (b: BookInput, execution?: Execution) => Object.values(compute(b, { execution }).sheets.at(-1)!.cells).map((c) => c.value)
const REF = { error: 'REF' }

describe('GETPIVOTDATA', () => {
  it.each(['closures', 'eval'] as Execution[])('reads the value field by caption or source name, items as labels, the grand total, and #REF! for what is not shown (%s)', (execution) => {
    const b = withFormulas([
      'GETPIVOTDATA("Sum of Amount",Pivot!A3)',
      'GETPIVOTDATA("amount",Pivot!B5,"REGION","south")',
      'GETPIVOTDATA("Amount",Pivot!A8,"Region",Data!A5)',
      'GETPIVOTDATA("Amount",Pivot!A3,"Region","Nowhere")',
      'GETPIVOTDATA("Amount",Pivot!A3,"Month","Jan")',
      'GETPIVOTDATA("Amount",Pivot!C3)',
      'GETPIVOTDATA("Count of Amount",Pivot!A3)',
      'GETPIVOTDATA("Amount",Pivot!A3,"Region")',
      'GETPIVOTDATA(Pivot!B3,Pivot!A3)',
      'GETPIVOTDATA("Amount","Pivot!A3")',
      'GETPIVOTDATA("Amount",Pivot!A3,"Region","North")/GETPIVOTDATA("Amount",Pivot!A3)',
      'GETPIVOTDATA(TEXT(1,"0")&"","Pivot!A3")',
    ])
    expect(values(b, execution)).toEqual([360, 80, 120, REF, REF, REF, REF, REF, REF, REF, 0.166666666666667, REF])
  })

  it('a PivotTable laid out before its last change (a stale cache) is #REF!', () => {
    const b = withFormulas(['GETPIVOTDATA("Amount",Pivot!A3)'], (x) => (x.pivotCaches![0].stale = true))
    expect(values(b)).toEqual([REF])
  })

  it('reads the layout the PivotTable has: its lines in another order over the same cells give each region the cell of its line', () => {
    const f = ['GETPIVOTDATA("Amount",Pivot!A3,"Region","North")', 'GETPIVOTDATA("Amount",Pivot!A3,"Region","West")']
    expect(values(withFormulas(f))).toEqual([60, 120])
    const swapped = withFormulas(f, (x) => {
      const lines = x.sheets[1].pivotTables![0].rowLines!
      ;[lines[0], lines[3]] = [lines[3], lines[0]]
    })
    expect(values(swapped)).toEqual([120, 60])
  })

  it("a collapsed item's line holds its total, in the tabular form too, where the inner field shows no line under it", () => {
    const f = ['GETPIVOTDATA("Amount",Pivot!A3,"Region","South")', 'GETPIVOTDATA("Amount",Pivot!A3,"Region","South","Month","Jan")']
    const collapsed = withFormulas(f, (x) => {
      const t = x.sheets[1].pivotTables![0]
      t.rows = [0, 2]
      t.layout = 'tabular'
      t.fields[0].items = t.fields[0].items!.map((i) => ({ ...i, collapsed: true }))
      t.fields[2] = { items: [{ item: 0 }, { item: 1 }] }
    })
    expect(values(collapsed)).toEqual([80, REF])
    // not collapsed, the tabular form shows no subtotal on the item's line
    const shown = withFormulas(f, (x) => {
      const t = x.sheets[1].pivotTables![0]
      t.rows = [0, 2]
      t.layout = 'tabular'
      t.fields[2] = { items: [{ item: 0 }, { item: 1 }] }
    })
    expect(values(shown)).toEqual([REF, REF])
  })

  it('a field of the Filters area by the item it shows, its cells part of the PivotTable', () => {
    const page = (item?: number) =>
      withFormulas(['GETPIVOTDATA("Amount",Pivot!A1,"Month","Jan")', 'GETPIVOTDATA("Amount",Pivot!B1)'], (x) => {
        const t = x.sheets[1].pivotTables![0]
        t.pages = [{ field: 2, ...(item === undefined ? {} : { item }) }]
        t.location.pageRows = 1
        t.fields[2] = { items: [{ item: 0 }, { item: 1 }] }
      })
    expect(values(page(0))).toEqual([360, 360])
    expect(values(page())).toEqual([REF, 360])
    expect(values(page(1))).toEqual([REF, 360])
  })
})

describe.each(['closures', 'eval'] as Execution[])('GETPIVOTDATA in a session, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    it('computes again when a refresh lays the PivotTable out anew, and back with undo', async () => {
      const s = await open(withFormulas(['GETPIVOTDATA("Amount",Pivot!A3,"Region","East")', 'GETPIVOTDATA("Amount",Pivot!A3)']), { execution })
      try {
        const at = async () => Object.values(await s.getCells('G', 'A1:A2')).map((c) => c.value)
        const refresh = async () => expect((await s.editStructure({ op: 'refreshPivots' })).ok).toBe(true)
        expect(await at()).toEqual([100, 360])
        // an East amount edited: the PivotTable keeps its cells until it is refreshed
        await s.setCells([{ sheet: 'Data', cell: 'B4', input: { value: 7 } }])
        expect(await at()).toEqual([100, 360])
        await refresh()
        expect(await at()).toEqual([77, 337])
        await s.setCells([{ sheet: 'Data', cell: 'B9', input: { value: 1 } }])
        await refresh()
        expect(await at()).toEqual([77, 258])
        for (let k = 0; k < 4; k++) await s.undo()
        expect(await at()).toEqual([100, 360])
        expect(await s.getHash()).toBe(compute(await s.getBook(), { execution }).hash)
      } finally {
        await s.close()
      }
    })
  })
})
