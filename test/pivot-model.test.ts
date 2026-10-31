/**
 * PivotTables in the workbook (`BookInput.pivotCaches`, `SheetInput.pivotTables` and the rest): the check of their
 * input, the core's of what it reads and of what refers to what, the full one of their shape; and a sheet renamed in a
 * session, which the caches and slicers follow.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput } from '../src/index.js'
import { writeXlsx } from '../src/xlsx.js'
import { pivotBook } from './pivot-books.js'
import { checkPivotShapes } from '../src/pivot/check.js'
import { setPivotCheck } from '../src/pivot/hook.js'

/** What `check` says of the book after `change`: 'accepted', or its refusal. */
function refusal(change: (b: BookInput) => void, check: (b: BookInput) => unknown = compute): string {
  const b = structuredClone(pivotBook())
  change(b)
  try {
    check(b)
    return 'accepted'
  } catch (e) {
    return (e as Error).message
  }
}

/** The full check of their shape, which the xlsx subpaths and `xlcalc/pivot` make. */
const full = (b: BookInput) => {
  checkPivotShapes(b)
  compute(b)
}

describe('PivotTables in the workbook', () => {
  it('a workbook with a PivotTable, its cache and a slicer is taken whole, and goes through a session unchanged but for the areas worked out', () => {
    expect(() => compute(pivotBook())).not.toThrow()
    const s = createSession(pivotBook())
    const want = pivotBook()
    want.sheets[1].pivotTables![0].location.areas = { table: 'A3:B8', columnHeaders: 'A3:B3', rowLabels: 'A4:A8', data: 'B4:B8' }
    expect(s.getBook()).toEqual(want)
  })

  it('the core refuses what it reads and what does not refer: a cache it does not have, fields not one for each of the cache, names twice, a slicer of another cache', () => {
    expect(refusal((b) => (b.sheets[1].pivotTables![0].cache = 2))).toContain('book.sheets[1].pivotTables[0].cache: expected the id of a pivot cache of the workbook, got number 2')
    expect(refusal((b) => b.sheets[1].pivotTables![0].fields.pop())).toContain("expected a field for each of the cache's 3 fields, got 2")
    expect(refusal((b) => (b.sheets[1].pivotTables![0].rows = [3]))).toContain("book.sheets[1].pivotTables[0].rows[0]: expected the index of a field of the PivotTable's cache, 0 to 2, got 3")
    expect(refusal((b) => (b.sheets[1].pivotTables![0].rowLines![0].items = [-1]))).toContain('rowLines[0].items[0]: expected an integer from 0')
    expect(refusal((b) => b.pivotCaches!.push({ ...b.pivotCaches![0] }))).toContain('book.pivotCaches[1].id: duplicate pivot cache id 1')
    expect(refusal((b) => b.sheets[1].pivotTables!.push(b.sheets[1].pivotTables![0]))).toContain('duplicate PivotTable name "PivotTable1"')
    expect(refusal((b) => (b.sheets[1].slicers![0].cache = 'Slicer_Month'))).toContain('book.sheets[1].slicers[0].cache: expected the name of a cache of the workbook, got "Slicer_Month"')
    expect(refusal((b) => (b.slicerCaches![0].pivots[0].name = 'PivotTable2'))).toContain('book.slicerCaches[0].pivots[0]: expected a PivotTable of the workbook')
    expect(refusal((b) => (b.pivotCaches![0].fields[0].items![0] = { date: 'soon' }))).toContain('book.pivotCaches[0].fields[0].items[0].date: invalid date "soon"')
  })

  it('the rest of their shape is checked in full by the xlsx subpaths, and by every check once xlcalc/pivot registers it', async () => {
    const colour = (b: BookInput) => ((b.pivotCaches![0] as unknown as Record<string, unknown>).colour = 1)
    expect(refusal(colour)).toBe('accepted')
    expect(refusal(colour, full)).toContain('book.pivotCaches[0]: unknown field "colour"')
    const book = structuredClone(pivotBook())
    colour(book)
    await expect(writeXlsx({ book })).rejects.toThrow('snapshot.book.pivotCaches[0]: unknown field "colour"')
    expect(refusal((b) => (b.pivotCaches![0].records![0][1] = { date: 'soon' }), full)).toContain('book.pivotCaches[0].records[0][1].date: invalid date "soon"')
    expect(refusal((b) => (b.pivotCaches![0].source = { type: 'cells' } as never), full)).toContain('book.pivotCaches[0].source.type: expected one of "range"')
    expect(refusal((b) => (b.sheets[1].pivotTables![0].values![0].showAs = 'percentOfGrand' as never), full)).toContain('values[0].showAs: expected one of')
    expect(refusal((b) => (b.sheets[1].conditionalFormats = [{ type: 'expression', ranges: ['B4:B7'], formula: 'B4>50', pivot: { pivot: 'PivotTable1', areas: [{ type: 'data', colour: 1 } as never] } }]), full)).toContain('pivot.areas[0]: unknown field "colour"')
    setPivotCheck(checkPivotShapes)
    try {
      expect(refusal(colour)).toContain('book.pivotCaches[0]: unknown field "colour"')
    } finally {
      setPivotCheck(null as never)
    }
  })

  it('a sheet renamed gives its new name to the source of a cache and to the PivotTables of a slicer', async () => {
    const s = createSession(pivotBook())
    s.editStructure({ op: 'renameSheet', sheet: 'Data', name: 'Raw' })
    s.editStructure({ op: 'renameSheet', sheet: 'Pivot', name: 'Report' })
    const b = s.getBook()
    expect(b.pivotCaches![0].source).toEqual({ type: 'range', sheet: 'Raw', range: 'A1:C9' })
    expect(b.slicerCaches![0].pivots).toEqual([{ sheet: 'Report', name: 'PivotTable1' }])
    await expect(writeXlsx(s.getSnapshot())).resolves.toBeDefined()
    s.undo()
    expect(s.getBook().slicerCaches![0].pivots).toEqual([{ sheet: 'Pivot', name: 'PivotTable1' }])
  })
})
