/**
 * Operations on PivotTables with `xlcalc/pivot` loaded: a new PivotTable is refreshed from its source and laid out, an
 * edit of its definition lays it out again from the cache, and a formula over it (`GETPIVOTDATA`) follows; a new or
 * moved PivotTable laid over data is refused. In every kind of session and both execution modes, with undo and redo.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, registerPivot, type BookInput } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)
registerPivot(pivotEngine)

const clock = { now: '2026-10-31T12:00:00Z', seed: 1 }

const book = (): BookInput => ({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'Region' }, B1: { value: 'Product' }, C1: { value: 'Amount' }, A2: { value: 'North' }, B2: { value: 'Apple' }, C2: { value: 10 }, A3: { value: 'South' }, B3: { value: 'Pear' }, C3: { value: 20 }, A4: { value: 'North' }, B4: { value: 'Pear' }, C4: { value: 5 } } },
    { name: 'P', cells: { F1: { formula: 'GETPIVOTDATA("Amount",$A$3,"Region","North")' } } },
  ],
})

const add = { op: 'addPivotTable', sheet: 'P', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:C4' }, definition: { rows: [0], values: [{ field: 2 }] } } as const

describe('addPivotTable and editPivotTable with xlcalc/pivot', () => {
  it('lays the new PivotTable out from its source, not stale', () => {
    const r = editStructure(book(), add)
    if (!r.ok) throw new Error(r.message)
    expect(r.stale).toBeUndefined()
    expect(r.book.pivotCaches![0].stale).toBeUndefined()
    const p = r.book.sheets[1].pivotTables![0]
    expect(p.location.range).toBe('A3:B6')
    const cells = r.book.sheets[1].cells
    expect(['A3', 'B3', 'A4', 'B4', 'A5', 'B5', 'A6', 'B6'].map((a) => cells[a]?.value)).toEqual(['Row Labels', 'Sum of Amount', 'North', 15, 'South', 20, 'Grand Total', 35])
  })

  it('refuses a new or moved PivotTable laid over data', () => {
    const over = { ...book() }
    over.sheets = [over.sheets[0], { ...over.sheets[1], cells: { ...over.sheets[1].cells, B30: { value: 'x' } } }]
    expect(editStructure(over, { ...add, cell: 'A29' })).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining("there's already data") })
    const r = editStructure(over, add)
    if (!r.ok) throw new Error(r.message)
    expect(editStructure(r.book, { op: 'movePivotTable', sheet: 'P', pivot: 'PivotTable1', to: { sheet: 'P', cell: 'A28' } })).toMatchObject({ ok: false, code: 'pivot' })
  })
})

describe.each(KINDS)('PivotTables laid out in %s', (kind) => {
  it.each(['closures', 'eval'] as const)('a field moved lays the cells out again and GETPIVOTDATA follows, with undo and redo (execution %s)', async (execution) => {
    const s = await (await opener(kind))(book(), { clock, execution })
    await s.editStructure(add)
    expect(await s.getValues('P', 'F1')).toEqual([[15]])
    // Product on the columns under the same numbers: the cells move, the formula follows the layout
    const r = await s.editStructure({ op: 'editPivotTable', sheet: 'P', pivot: 'PivotTable1', set: { columns: [1] } })
    expect(r).toMatchObject({ ok: true })
    expect(await s.getValues('P', 'A4:D5')).toEqual([['Row Labels', 'Apple', 'Pear', 'Grand Total'], ['North', 10, 5, 15]])
    expect(await s.getValues('P', 'F1')).toEqual([[15]])
    await s.editStructure({ op: 'editPivotTable', sheet: 'P', pivot: 'PivotTable1', set: { values: [{ field: 2, function: 'count' }] } })
    // the value field by its source field's name: now the count of North's rows
    expect(await s.getValues('P', 'F1')).toEqual([[2]])
    await s.undo()
    expect(await s.getValues('P', 'F1')).toEqual([[15]])
    await s.undo()
    expect(await s.getValues('P', 'A3:B6')).toEqual([['Row Labels', 'Sum of Amount'], ['North', 15], ['South', 20], ['Grand Total', 35]])
    await s.redo()
    expect(await s.getValues('P', 'D5')).toEqual([[15]])
    await s.close()
  })
})
