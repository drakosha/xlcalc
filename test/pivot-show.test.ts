/**
 * Show Values As over a PivotTable refreshed from its source: the calculations at the lines of an axis and its grand
 * total, a base field on no axis or none at all (#N/A), and a base item kept by its label through a refresh that drops
 * an item before it, a base field gone with its header. Against Excel: the sets `show` and `show-items`
 * (`test/pivot-excel.test.ts`).
 */
import { describe, expect, it } from 'vitest'
import { editStructure, registerPivot, type BookInput, type CellValue, type PivotDataField } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'

registerPivot(pivotEngine)

const rows = (regions: string[], amounts: number[]) =>
  Object.fromEntries([['A1', { value: 'Region' }], ['B1', { value: 'Amount' }], ...regions.flatMap((r, i) => [[`A${i + 2}`, { value: r }], [`B${i + 2}`, { value: amounts[i] }]])])

/** Region on rows (East, North, South, West ascending), Amount summed and shown as `d` says. */
function book(d: Omit<PivotDataField, 'field'>, retainItems?: number): BookInput {
  return {
    sheets: [
      { name: 'D', cells: rows(['North', 'South', 'East', 'West', 'North'], [10, 20, 5, 40, 1]) },
      { name: 'P', cells: {}, pivotTables: [{ name: 'PT1', cache: 1, location: { range: 'A3:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: [{}, {}], rows: [0], values: [{ field: 1, ...d }] }] },
    ],
    pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:B6' }, fields: [{ name: 'Region' }, { name: 'Amount' }], ...(retainItems === undefined ? {} : { retainItems }) }],
  }
}

const REFRESH = { op: 'refreshPivots' } as const

/** The values of the regions' lines and the grand total, B4:B8. */
function shown(b: BookInput): CellValue[] {
  const r = editStructure(b, REFRESH)
  if (!r.ok) throw new Error(r.message)
  return [4, 5, 6, 7, 8].map((n) => {
    const c = r.book.sheets[1].cells[`B${n}`]
    return c?.formula?.startsWith('#') ? c.formula : (c?.value ?? null)
  }) as CellValue[]
}

describe('Show Values As', () => {
  it('calculates each line of the rows and the grand total', () => {
    expect(shown(book({ showAs: 'percentOfTotal' }))).toEqual([5 / 76, 11 / 76, 20 / 76, 40 / 76, 1])
    expect(shown(book({ showAs: 'runTotal', baseField: 0 }))).toEqual([5, 16, 36, 76, null])
    expect(shown(book({ showAs: 'percentOfRunningTotal', baseField: 0 }))).toEqual([5 / 76, 16 / 76, 36 / 76, 1, null])
    expect(shown(book({ showAs: 'rankDescending', baseField: 0 }))).toEqual([4, 3, 2, 1, null])
    expect(shown(book({ showAs: 'difference', baseField: 0, baseItem: 'previous' }))).toEqual([null, 6, 9, 20, null])
    expect(shown(book({ showAs: 'percent', baseField: 0, baseItem: 1 }))).toEqual([5 / 11, 1, 20 / 11, 40 / 11, null])
    expect(shown(book({ showAs: 'index', baseField: 0 }))).toEqual([1, 1, 1, 1, 1])
  })

  it('is #N/A by a base field on no axis, without a base field, or by a base item hidden', () => {
    const na = ['#N/A', '#N/A', '#N/A', '#N/A', '#N/A']
    expect(shown(book({ showAs: 'difference', baseField: 1, baseItem: 0 }))).toEqual(na)
    expect(shown(book({ showAs: 'runTotal' }))).toEqual(na)
    // North, the base item, hidden: its regions and the grand total, none below
    const r = editStructure(book({ showAs: 'percent', baseField: 0, baseItem: 1 }), REFRESH)
    if (!r.ok) throw new Error(r.message)
    r.book.sheets[1].pivotTables![0].fields[0].items![1].hidden = true
    expect(shown(r.book)).toEqual(['#N/A', '#N/A', '#N/A', '#N/A', null])
  })

  it('keeps its base item by label through a refresh that drops an item before it, and loses a base field gone', () => {
    const r = editStructure(book({ showAs: 'percent', baseField: 0, baseItem: 2 }, 0), REFRESH)
    if (!r.ok) throw new Error(r.message)
    // East gone from the source, and with retainItems 0 from the field: South moves from the third item to the second
    const edited = structuredClone(r.book)
    edited.sheets[0].cells.A4 = { value: 'North' }
    expect(shown(edited)).toEqual([16 / 20, 1, 40 / 20, null, null])
    edited.sheets[0].cells.A1 = { value: 'Area' }
    const gone = editStructure(edited, REFRESH)
    expect(gone.ok && gone.book.sheets[1].pivotTables![0].values).toEqual([{ field: 1, showAs: 'percent' }])
  })
})
