/**
 * A session refuses an entry into a PivotTable's area, as Excel refuses one: `setCells`, `replaceCells`, `clearCells`
 * and `applyChanges` throw `PivotInputError`, `setCellText`, `setCellTexts`, `fillCellText` and `pasteText` answer `pivot`,
 * each with Excel's alert (`reason`: a value or a paste changes a part, a cleared cell affects it, a formula over a label),
 * the Filters area too (not the blank row under it); a cell's style and number format may change, and a cell outside the area takes anything. In
 * every kind of session and both execution modes; nothing is written, so there is no step to undo.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, PivotInputError, type BookInput, type PivotTable } from '../src/index.js'
import { pivotEntry } from '../src/pivot/entry.js'
import { setPivotEntry, setPivotRefresh } from '../src/pivot/hook.js'
import { refreshPivots } from '../src/pivot/engine.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-31T12:00:00Z', seed: 1 }

const pivot: PivotTable = {
  name: 'PT',
  cache: 1,
  location: { range: 'A3:B5', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, pageRows: 1, pageColumns: 1 },
  fields: [{ items: [{ item: 0 }] }, {}, { items: [{ item: 0 }] }],
  rows: [0],
  pages: [{ field: 2 }],
  values: [{ field: 1, name: 'Sum of N' }],
  rowLines: [{ items: [0] }, { type: 'grand', items: [0] }],
  columnLines: [{ items: [] }],
}

const book = (): BookInput => ({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'K' }, B1: { value: 'N' }, C1: { value: 'M' }, A2: { value: 'a' }, B2: { value: 2 }, C2: { value: 'x' } } },
    {
      name: 'P',
      cells: { A1: { value: 'M' }, B1: { value: '(All)' }, A3: { value: 'Row Labels' }, B3: { value: 'Sum of N' }, A4: { value: 'a' }, B4: { value: 2 }, A5: { value: 'Grand Total' }, B5: { value: 2 } },
      pivotTables: [pivot],
    },
  ],
  pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'Data', range: 'A1:C2' }, fields: [{ name: 'K', items: ['a'] }, { name: 'N' }, { name: 'M', items: ['x'] }], records: [[0, 2, 0]] }],
})

const refusal = { ok: false, code: 'pivot', message: "P!B4 is part of the PivotTable PT: we can't change this part of the PivotTable", reason: 'part' }

/** The error an edit throws over a PivotTable's cell: a `PivotInputError`, an `InputError` too, with Excel's alert and the cell. */
async function thrown(p: Promise<unknown>, reason: string, cell: string) {
  const e = await p.then(() => null, (x: unknown) => x)
  expect(e).toBeInstanceOf(PivotInputError)
  expect(e).toBeInstanceOf(InputError)
  expect(e).toMatchObject({ name: 'PivotInputError', path: 'edits', reason, sheet: 'P', cell })
}

describe.each(KINDS)('an entry into a PivotTable in %s', (kind) => {
  it.each(['closures', 'eval'] as const)('is refused, the Filters area too; looks may change, cells outside take anything (execution %s)', async (execution) => {
    const s = await (await opener(kind))(book(), { clock, execution })
    await thrown(s.setCells([{ sheet: 'P', cell: 'B4', input: { value: 3 } }]), 'part', 'B4')
    await expect(s.setCells([{ sheet: 'P', cell: 'B1', input: null }])).rejects.toThrow(/part of the PivotTable PT/)
    await thrown(s.setCells([{ sheet: 'P', cell: 'B1', input: null }]), 'affect', 'B1')
    await thrown(s.setCells([{ sheet: 'p', cell: 'C4', input: { value: 1 } }, { sheet: 'p', cell: 'B4', input: { formula: '1' } }]), 'part', 'B4')
    await thrown(s.replaceCells([{ sheet: 'P', cell: 'A5', input: { value: 'All' } }]), 'part', 'A5')
    await thrown(s.clearCells([{ sheet: 'P', range: 'A4:C4' }], 'contents'), 'affect', 'A4')
    await thrown(s.applyChanges([{ op: 'setCells', edits: [{ sheet: 'P', cell: 'A4', input: { formula: '1+1' } }] }]), 'formula', 'A4')
    await thrown(s.validateEntry('P', 'B4', { value: 1 }), 'part', 'B4')
    expect(await s.setCellText('P', 'B4', '3')).toMatchObject(refusal)
    expect(await s.previewCellText('P', 'B4', '3')).toMatchObject(refusal)
    // a paste affects it, one cell's text too, as Excel's alert says
    const affect = { ok: false, code: 'pivot', reason: 'affect', message: expect.stringContaining('will affect a PivotTable') }
    expect(await s.pasteText('P', 'A4', '1\t2')).toMatchObject({ ...affect, sheet: 'P', cell: 'A4', message: expect.stringContaining('P!A4') })
    expect(await s.previewPasteText('P', 'B4', '1')).toMatchObject({ ...affect, sheet: 'P', cell: 'B4' })
    expect(await s.fillCellText('P', 'C4:C5', '1')).toMatchObject({ ok: true })
    expect(await s.fillCellText('P', 'B5:C5', '1')).toMatchObject({ ok: false, code: 'pivot', reason: 'part' })
    expect(await s.setCellTexts([{ sheet: 'P', cell: 'C1', text: 'x' }, { sheet: 'P', cell: 'A3', text: 'x' }])).toMatchObject({ ...affect, sheet: 'P', cell: 'A3' })
    expect(await s.setCellTexts([{ sheet: 'P', cell: 'B4', text: '1' }])).toMatchObject({ ...refusal, sheet: 'P', cell: 'B4' })
    // a look, and a cell beyond the area
    await s.setCells([{ sheet: 'P', cell: 'B4', input: { numberFormat: '0.00' } }, { sheet: 'P', cell: 'D4', input: { value: 1 } }])
    await s.clearCells([{ sheet: 'P', range: 'B4' }], 'formats')
    expect(await s.getValues('P', 'B4:D5')).toEqual([[2, 1, 1], [2, 1, null]])
    // the blank row between the Filters area and the table takes anything, as in Excel
    await s.setCellText('P', 'A2', 'x')
    expect(await s.getValues('P', 'A2')).toEqual([['x']])
    // its table and Filters area cleared whole delete it, the blank row between keeping what it holds, as in Excel
    expect(await s.clearCells([{ sheet: 'P', range: 'A1:B1' }, { sheet: 'P', range: 'A3:B5' }], 'contents')).toMatchObject({ layout: ['P'] })
    expect((await s.getBook()).sheets[1].pivotTables).toBeUndefined()
    expect(await s.getValues('P', 'A1:B5')).toEqual([[null, null], ['x', null], [null, null], [null, null], [null, null]])
    await s.close()
  })
})

describe.each(KINDS.slice(0, 2))('an entry into a PivotTable with xlcalc/pivot in %s', (kind) => {
  it('names its part through editPivotTable, keeps its same name, and refuses as Excel does', async () => {
    setPivotEntry(pivotEntry)
    setPivotRefresh(refreshPivots)
    try {
      const s = await (await opener(kind))(book(), { clock })
      expect(await s.setCellText('P', 'B4', '3')).toMatchObject(refusal)
      expect(await s.setCellText('P', 'A4', 'A')).toMatchObject({ ok: true, changes: [] })
      const taken = await s.setCellText('P', 'B3', 'n')
      expect(taken).toMatchObject({ ok: false, code: 'pivot', message: 'PivotTable field name already exists: "n"' })
      expect(taken).not.toHaveProperty('reason')
      expect(await s.setCellText('P', 'A4', '=1')).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('cannot enter a formula'), reason: 'formula' })
      await expect(s.setCells([{ sheet: 'P', cell: 'A4', input: null }])).rejects.toThrow(/affect a PivotTable/)
      await thrown(s.setCells([{ sheet: 'P', cell: 'A4', input: null }]), 'affect', 'A4')
      await thrown(s.setCells([{ sheet: 'P', cell: 'A4', input: { formula: '1' } }]), 'formula', 'A4')
      const item = await s.setCells([{ sheet: 'P', cell: 'B1', input: { value: 'b' } }]).then(() => null, (e: unknown) => e)
      expect(item).toBeInstanceOf(InputError)
      expect(item).not.toBeInstanceOf(PivotInputError)
      expect((item as Error).message).toContain('"b" is not an item of this field')
      // one text of setCellTexts names its part as setCellText does; an entry into a selection, the active cell's alone
      expect(await s.setCellTexts([{ sheet: 'P', cell: 'B3', text: 'Total' }])).toMatchObject({ ok: true })
      expect(await s.getValues('P', 'B3')).toEqual([['Total']])
      expect(await s.setCellTexts([{ sheet: 'P', cell: 'B4', text: '1' }])).toMatchObject({ ok: false, code: 'pivot', reason: 'part', sheet: 'P', cell: 'B4' })
      expect(await s.fillCellText('P', 'A4:B4', 'y')).toMatchObject({ ok: true })
      expect(await s.getValues('P', 'A4:B4')).toEqual([['y', 2]])
      expect(await s.fillCellText('P', 'A4:B4', '1', { active: 'B4' })).toMatchObject({ ok: false, code: 'pivot', reason: 'part', sheet: 'P', cell: 'B4' })
      expect(await s.fillCellText('P', 'A4:A5', '=1')).toMatchObject({ ok: false, code: 'pivot', reason: 'formula', sheet: 'P', cell: 'A4' })
      await s.undo()
      await s.undo()
      // a new name is an edit of the definition, which the session makes with editPivotTable: the item's caption, laid out again
      expect(await s.setCellText('P', 'A4', 'z')).toMatchObject({ ok: true })
      expect(await s.getValues('P', 'A3:B5')).toEqual([['Row Labels', 'Sum of N'], ['z', 2], ['Grand Total', 2]])
      expect((await s.getBook()).sheets[1].pivotTables![0].fields[0].items).toEqual([{ item: 0, caption: 'z' }])
      // its contents cleared whole, with other cells, delete it in one step, as Excel does; part of it is refused first
      await s.setCells([{ sheet: 'P', cell: 'D9', input: { value: 1 } }])
      await thrown(s.clearCells([{ sheet: 'P', range: 'A3:B5' }, { sheet: 'P', range: 'D9' }], 'contents'), 'affect', 'A3')
      expect(await s.clearCells([{ sheet: 'P', range: 'A1:B1' }, { sheet: 'P', range: 'A3:D9' }], 'contents')).toMatchObject({ layout: ['P'], changes: expect.arrayContaining([{ sheet: 'P', cell: 'A4', text: '', value: null }, { sheet: 'P', cell: 'D9', text: '', value: null }]) })
      expect((await s.getBook()).sheets[1].pivotTables).toBeUndefined()
      expect(await s.getValues('P', 'A1:D9')).toEqual(Array.from({ length: 9 }, () => [null, null, null, null]))
      await s.undo()
      expect((await s.getBook()).sheets[1].pivotTables!.map((t) => t.name)).toEqual(['PT'])
      expect(await s.getValues('P', 'D9')).toEqual([[1]])
      await s.close()
    } finally {
      setPivotEntry(() => null)
      setPivotRefresh(null)
    }
  })
})
