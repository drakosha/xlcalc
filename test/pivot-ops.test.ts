/**
 * Operations on PivotTables without `xlcalc/pivot` (`editStructure`): a PivotTable created over a range, a table, a
 * name or a cache, its definition changed, renamed, moved, deleted; the operations of the structure moving it with its
 * cells and refusing to cut it; a change it cannot lay out marks its cache stale, which a file asks Excel to refresh on
 * opening. In every kind of session and both execution modes, with undo and redo.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, InputError, type BookInput, type PivotTable, type StructureOp } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-31T12:00:00Z', seed: 1 }

const data = (): BookInput => ({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'Region' }, B1: { value: 'Amount' }, C1: { value: 'Amount' }, A2: { value: 'North' }, B2: { value: 10 }, C2: { value: 1 }, A3: { value: 'South' }, B3: { value: 20 }, C3: { value: 2 }, A4: { value: 'North' }, B4: { value: 5 }, C4: { value: 3 }, E1: { value: 'K' }, F1: { value: 'V' }, E2: { value: 'a' }, F2: { value: 1 } }, tables: [{ name: 'Sales', range: 'E1:F3', columns: [{ name: 'K' }, { name: 'V' }] }] },
    { name: 'P', cells: { H1: { value: 'note' } } },
  ],
  names: [{ name: 'Src', ref: 'Data!$A$1:$C$4' }],
})

const add = (more: Partial<Extract<StructureOp, { op: 'addPivotTable' }>> = {}): StructureOp => ({ op: 'addPivotTable', sheet: 'P', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:C4' }, ...more })

const pivots = (b: BookInput, s = 1) => b.sheets[s].pivotTables ?? []

describe('addPivotTable without xlcalc/pivot', () => {
  it('makes a cache of the header row, the PivotTable at its place, its cache stale', () => {
    const r = editStructure(data(), add())
    expect(r).toMatchObject({ ok: true, stale: [{ sheet: 'P', name: 'PivotTable1' }] })
    if (!r.ok) return
    // a repeated header gets 2, as Excel names it
    expect(r.book.pivotCaches).toEqual([{ id: 1, source: { type: 'range', sheet: 'Data', range: 'A1:C4' }, fields: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Amount2' }], stale: true }])
    expect(pivots(r.book)).toEqual([{ name: 'PivotTable1', cache: 1, location: { range: 'A3:C20', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, fields: [{}, {}, {}], indent: 0, multipleFieldFilters: false, style: { name: 'PivotStyleLight16', rowHeaders: true, columnHeaders: true, lastColumn: true } }])
  })

  it('takes a table, a name, a cache, a definition and a new sheet', () => {
    const t = editStructure(data(), add({ source: { type: 'table', table: 'sales' }, name: 'T' }))
    expect(t.ok && t.book.pivotCaches![0].fields.map((f) => f.name)).toEqual(['K', 'V'])
    const n = editStructure(data(), add({ source: { type: 'name', name: 'src' } }))
    expect(n.ok && n.book.pivotCaches![0].fields.map((f) => f.name)).toEqual(['Region', 'Amount', 'Amount2'])
    const one = editStructure(data(), add({ definition: { rows: [0], values: [{ field: 1 }, { field: 1 }, { field: 2, function: 'count' }] } }))
    if (!one.ok) throw new Error(one.message)
    expect(pivots(one.book)[0]).toMatchObject({ rows: [0], values: [{ field: 1, name: 'Sum of Amount' }, { field: 1, name: 'Sum of Amount2' }, { field: 2, function: 'count', name: 'Count of Amount2' }] })
    const two = editStructure(one.book, add({ source: { cache: 1 }, sheet: 'New', cell: undefined, newSheet: { before: 'P' } }))
    expect(two.ok && [two.book.sheets.map((s) => s.name), pivots(two.book)[0], two.book.pivotCaches!.length]).toMatchObject([['Data', 'New', 'P'], { name: 'PivotTable2', cache: 1, location: { range: 'A3:C20' } }, 1])
  })

  it('refuses a place over data, a table or a PivotTable, a name taken, an empty header; a wrong shape throws', () => {
    expect(editStructure(data(), add({ cell: 'G1' }))).toMatchObject({ ok: false, code: 'pivot', message: "there's already data in P!G1:I18, where the PivotTable PivotTable1 would be laid out" })
    expect(editStructure(data(), add({ sheet: 'Data', cell: 'E3' }))).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('the table Sales') })
    const one = editStructure(data(), add())
    if (!one.ok) throw new Error(one.message)
    expect(editStructure(one.book, add({ cell: 'B10' }))).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('overlap the PivotTable PivotTable1') })
    expect(editStructure(one.book, add({ cell: 'E30', name: 'pivottable1' }))).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('already exists') })
    expect(editStructure(data(), add({ source: { type: 'range', sheet: 'Data', range: 'B1:D4' } }))).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('field name is not valid') })
    expect(editStructure(data(), add({ source: { type: 'range', sheet: 'Data', range: 'A1:C1' } }))).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('at least one row') })
    expect(() => editStructure(data(), add({ source: { cache: 9 } }))).toThrow(InputError)
    expect(() => editStructure(data(), add({ definition: { rows: [0], columns: [0] } }))).toThrow('a field lies on one axis')
    expect(() => editStructure(data(), add({ definition: { name: 'x' } as never }))).toThrow(InputError)
  })
})

/** A book with a PivotTable of Region by Amount on P!A3 (stale, as made). */
function made(): BookInput {
  const r = editStructure(data(), add({ definition: { rows: [0], values: [{ field: 1 }] } }))
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe('editPivotTable, renamePivotTable, movePivotTable, deletePivotTable without xlcalc/pivot', () => {
  it('patches the definition: null removes, fields by index, Change Data Source keeps fields by name', () => {
    const r = editStructure(made(), { op: 'editPivotTable', sheet: 'P', pivot: 'pivottable1', set: { rows: null, columns: [0], layout: 'tabular', fields: [{ field: 0, subtotals: [] }] } })
    if (!r.ok) throw new Error(r.message)
    const p = pivots(r.book)[0]
    expect([p.rows, p.columns, p.layout, p.fields[0]]).toEqual([undefined, [0], 'tabular', { subtotals: [] }])
    const s = editStructure(r.book, { op: 'editPivotTable', sheet: 'P', pivot: 'PivotTable1', set: {}, source: { type: 'range', sheet: 'Data', range: 'B1:B4' } })
    if (!s.ok) throw new Error(s.message)
    expect([pivots(s.book)[0].columns, pivots(s.book)[0].values, s.book.pivotCaches![0].fields]).toEqual([[], [{ field: 0, name: 'Sum of Amount' }], [{ name: 'Amount' }]])
  })

  it('renames, keeping the names of rules and slicers; refuses a name taken on its sheet', () => {
    const b = made()
    const r = editStructure(b, { op: 'renamePivotTable', sheet: 'P', pivot: 'PivotTable1', name: 'Sales by region' })
    expect(r.ok && pivots(r.book)[0].name).toBe('Sales by region')
    const two = editStructure(b, add({ cell: 'E30', name: 'Other' }))
    if (!two.ok) throw new Error(two.message)
    expect(editStructure(two.book, { op: 'renamePivotTable', sheet: 'P', pivot: 'Other', name: 'PIVOTTABLE1' })).toMatchObject({ ok: false, code: 'pivot' })
  })

  it('moves with its cells, to another sheet too, and deletes with them', () => {
    const b = made()
    const sheets = [...b.sheets]
    sheets[1] = { ...sheets[1], cells: { ...sheets[1].cells, A3: { value: 'Row Labels' }, B4: { value: 15 } } }
    const m = editStructure({ ...b, sheets }, { op: 'movePivotTable', sheet: 'P', pivot: 'PivotTable1', to: { sheet: 'Data', cell: 'K2' } })
    if (!m.ok) throw new Error(m.message)
    expect([pivots(m.book), pivots(m.book, 0)[0].location.range, m.book.sheets[0].cells.K2, m.book.sheets[0].cells.L3]).toEqual([[], 'K2:M19', { value: 'Row Labels' }, { value: 15 }])
    const d = editStructure(m.book, { op: 'deletePivotTable', sheet: 'Data', pivot: 'PivotTable1' })
    expect(d.ok && [d.book.pivotCaches, d.book.sheets[0].cells.K2, d.book.sheets[0].pivotTables]).toEqual([undefined, undefined, undefined])
  })

  it('refreshPivots without xlcalc/pivot is refused', () => {
    expect(editStructure(made(), { op: 'refreshPivots' })).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('xlcalc/pivot') })
    expect(editStructure(made(), { op: 'refreshPivots', pivots: [{ sheet: 'P', name: 'Nope' }] })).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('Nope') })
  })
})

describe('operations of the structure over a PivotTable', () => {
  const op = (o: StructureOp) => editStructure(made(), o)
  const at = (r: ReturnType<typeof op>) => (r.ok ? r.book.sheets.map((s) => (s.pivotTables ?? []).map((p: PivotTable) => p.location.range)) : r.code)
  it.each<[string, StructureOp, unknown]>([
    ['rows inserted above move it', { op: 'insertRows', sheet: 'P', row: 1, count: 2 }, [[], ['A5:C22']]],
    ['rows inserted into it are refused', { op: 'insertRows', sheet: 'P', row: 5 }, 'pivot'],
    ['columns deleted before it move it', { op: 'insertColumns', sheet: 'P', column: 'A' }, [[], ['B3:D20']]],
    ['its rows deleted all delete it', { op: 'deleteRows', sheet: 'P', row: 2, count: 30 }, [[], []]],
    ['part of its rows deleted is refused', { op: 'deleteRows', sheet: 'P', row: 4 }, 'pivot'],
    ['all its cells cut move it', { op: 'moveRange', from: { sheet: 'P', range: 'A3:C20' }, to: { sheet: 'Data', cell: 'K1' } }, [['K1:M18'], []]],
    ['part of its cells cut is refused', { op: 'moveRange', from: { sheet: 'P', range: 'A3:B9' }, to: { sheet: 'P', cell: 'K1' } }, 'pivot'],
    ['a paste onto it is refused', { op: 'copyRange', from: { sheet: 'P', range: 'H1' }, to: { sheet: 'P', range: 'B5' } }, 'pivot'],
    ['a merge over it is refused', { op: 'mergeCells', sheet: 'P', range: 'C20:D21' }, 'pivot'],
    ['its sheet deleted deletes it', { op: 'deleteSheet', sheet: 'P' }, [[]]],
  ])('%s', (_, o, want) => expect(at(op(o))).toEqual(want))

  it('moves its source with the source cells, keeps it when they or their sheet are deleted, renames its sheet', () => {
    const src = (o: StructureOp) => {
      const r = op(o)
      return r.ok ? r.book.pivotCaches![0].source : r.code
    }
    expect(src({ op: 'insertRows', sheet: 'Data', row: 3, count: 2 })).toEqual({ type: 'range', sheet: 'Data', range: 'A1:C6' })
    expect(src({ op: 'insertColumns', sheet: 'Data', column: 'A' })).toEqual({ type: 'range', sheet: 'Data', range: 'B1:D4' })
    expect(src({ op: 'deleteColumns', sheet: 'Data', column: 'A', count: 3 })).toEqual({ type: 'range', sheet: 'Data', range: 'A1:C4' })
    expect(src({ op: 'renameSheet', sheet: 'Data', name: 'Rows' })).toEqual({ type: 'range', sheet: 'Rows', range: 'A1:C4' })
    expect(src({ op: 'deleteSheet', sheet: 'Data' })).toEqual({ type: 'range', sheet: 'Data', range: 'A1:C4' })
  })
})

describe('a slicer over a PivotTable whose sheet is deleted', () => {
  it('keeps its cache, its choice and the pivot cache, naming no PivotTable, as Excel does', () => {
    const b = made()
    const sheets = [b.sheets[0], b.sheets[1], { name: 'S', cells: {}, slicers: [{ name: 'Region', cache: 'Slicer_Region' }] }]
    const book: BookInput = { ...b, sheets, slicerCaches: [{ name: 'Slicer_Region', field: 'Region', pivots: [{ sheet: 'P', name: 'PivotTable1' }], cache: 1, items: [{ item: 0 }, { item: 1, selected: false }] }] }
    const r = editStructure(book, { op: 'deleteSheet', sheet: 'P' })
    if (!r.ok) throw new Error(r.message)
    expect([r.book.slicerCaches, r.book.sheets[1].slicers?.length, r.book.pivotCaches?.map((c) => c.id)]).toEqual([[{ name: 'Slicer_Region', field: 'Region', pivots: [], cache: 1, items: [{ item: 0 }, { item: 1, selected: false }] }], 1, [1]])
  })
})

describe('a stale cache in a file', () => {
  it('is written to be refreshed on opening, and read back as such', async () => {
    const w = await writeXlsx({ book: made() }, { values: 'omit' })
    const r = await readXlsx(w.data)
    expect(r.book.pivotCaches?.[0]).toMatchObject({ refreshOnLoad: true })
  })
})

describe.each(KINDS)('PivotTable operations in %s', (kind) => {
  it.each(['closures', 'eval'] as const)('add, edit, rename, move, delete with undo and redo (execution %s)', async (execution) => {
    const s = await (await opener(kind))(data(), { clock, execution })
    expect(await s.editStructure(add())).toMatchObject({ ok: true, stale: [{ sheet: 'P', name: 'PivotTable1' }] })
    await s.editStructure({ op: 'editPivotTable', sheet: 'P', pivot: 'PivotTable1', set: { rows: [0] } })
    await s.editStructure({ op: 'renamePivotTable', sheet: 'P', pivot: 'PivotTable1', name: 'R' })
    await s.editStructure({ op: 'movePivotTable', sheet: 'P', pivot: 'R', to: { sheet: 'P', cell: 'A30' } })
    const names = async () => (await s.getBook()).sheets[1].pivotTables?.map((p) => [p.name, p.location.range, p.rows])
    expect(await names()).toEqual([['R', 'A30:C47', [0]]])
    await s.editStructure({ op: 'deletePivotTable', sheet: 'P', pivot: 'R' })
    expect(await names()).toBeUndefined()
    await s.undo()
    expect(await names()).toEqual([['R', 'A30:C47', [0]]])
    await s.undo()
    expect(await names()).toEqual([['R', 'A3:C20', [0]]])
    await s.undo()
    expect(await names()).toEqual([['PivotTable1', 'A3:C20', [0]]])
    await s.redo()
    await s.redo()
    expect(await names()).toEqual([['R', 'A30:C47', [0]]])
    expect(await s.editStructure({ op: 'insertRows', sheet: 'P', row: 31 })).toMatchObject({ ok: false, code: 'pivot' })
    await s.close()
  })
})
