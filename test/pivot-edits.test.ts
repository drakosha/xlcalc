/**
 * A session refuses an entry into a PivotTable's area, as Excel refuses one: `setCells`, `replaceCells`, `clearCells`
 * and `applyChanges` throw `InputError`, `setCellText`, `setCellTexts`, `fillCellText` and `pasteText` answer `pivot`,
 * the Filters area too (not the blank row under it); a cell's style and number format may change, and a cell outside the area takes anything. In
 * every kind of session and both execution modes; nothing is written, so there is no step to undo.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type BookInput, type PivotTable } from '../src/index.js'
import { pivotEntry } from '../src/pivot/entry.js'
import { setPivotEntry } from '../src/pivot/hook.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-31T12:00:00Z', seed: 1 }

const pivot: PivotTable = {
  name: 'PT',
  cache: 1,
  location: { range: 'A3:B5', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, pageRows: 1, pageColumns: 1 },
  fields: [{ items: [{ item: 0 }] }, {}],
  rows: [0],
  pages: [{ field: 0 }],
  values: [{ field: 1, name: 'Sum of N' }],
  rowLines: [{ items: [0] }, { type: 'grand', items: [0] }],
  columnLines: [{ items: [] }],
}

const book = (): BookInput => ({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'K' }, B1: { value: 'N' }, A2: { value: 'a' }, B2: { value: 2 } } },
    {
      name: 'P',
      cells: { A1: { value: 'K' }, B1: { value: '(All)' }, A3: { value: 'Row Labels' }, B3: { value: 'Sum of N' }, A4: { value: 'a' }, B4: { value: 2 }, A5: { value: 'Grand Total' }, B5: { value: 2 } },
      pivotTables: [pivot],
    },
  ],
  pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'Data', range: 'A1:B2' }, fields: [{ name: 'K', items: ['a'] }, { name: 'N' }] }],
})

const refusal = { ok: false, code: 'pivot', message: "P!B4 is part of the PivotTable PT: we can't change this part of the PivotTable" }

describe.each(KINDS)('an entry into a PivotTable in %s', (kind) => {
  it.each(['closures', 'eval'] as const)('is refused, the Filters area too; looks may change, cells outside take anything (execution %s)', async (execution) => {
    const s = await (await opener(kind))(book(), { clock, execution })
    await expect(s.setCells([{ sheet: 'P', cell: 'B4', input: { value: 3 } }])).rejects.toThrow(InputError)
    await expect(s.setCells([{ sheet: 'P', cell: 'B1', input: null }])).rejects.toThrow(/part of the PivotTable PT/)
    await expect(s.replaceCells([{ sheet: 'P', cell: 'A5', input: { value: 'All' } }])).rejects.toThrow(InputError)
    await expect(s.clearCells([{ sheet: 'P', range: 'A4:C4' }], 'contents')).rejects.toThrow(InputError)
    await expect(s.applyChanges([{ op: 'setCells', edits: [{ sheet: 'P', cell: 'A4', input: { formula: '1+1' } }] }])).rejects.toThrow(InputError)
    expect(await s.setCellText('P', 'B4', '3')).toMatchObject(refusal)
    expect(await s.previewCellText('P', 'B4', '3')).toMatchObject(refusal)
    expect(await s.pasteText('P', 'A4', '1\t2')).toMatchObject({ ...refusal, sheet: 'P', cell: 'A4', message: expect.stringContaining('P!A4') })
    expect(await s.fillCellText('P', 'C4:C5', '1')).toMatchObject({ ok: true })
    expect(await s.fillCellText('P', 'B5:C5', '1')).toMatchObject({ ok: false, code: 'pivot' })
    expect(await s.setCellTexts([{ sheet: 'P', cell: 'C1', text: 'x' }, { sheet: 'P', cell: 'A3', text: 'x' }])).toMatchObject({ ok: false, code: 'pivot' })
    // a look, and a cell beyond the area
    await s.setCells([{ sheet: 'P', cell: 'B4', input: { numberFormat: '0.00' } }, { sheet: 'P', cell: 'D4', input: { value: 1 } }])
    await s.clearCells([{ sheet: 'P', range: 'B4' }], 'formats')
    expect(await s.getValues('P', 'B4:D5')).toEqual([[2, 1, 1], [2, 1, null]])
    // the blank row between the Filters area and the table takes anything, as in Excel
    await s.setCellText('P', 'A2', 'x')
    expect(await s.getValues('P', 'A2')).toEqual([['x']])
    await s.close()
  })
})

describe.each(KINDS.slice(0, 2))('an entry into a PivotTable with xlcalc/pivot in %s', (kind) => {
  it('names its part through editPivotTable, keeps its same name, and refuses as Excel does', async () => {
    setPivotEntry(pivotEntry)
    try {
      const s = await (await opener(kind))(book(), { clock })
      expect(await s.setCellText('P', 'B4', '3')).toMatchObject(refusal)
      expect(await s.setCellText('P', 'A4', 'A')).toMatchObject({ ok: true, changes: [] })
      expect(await s.setCellText('P', 'B3', 'n')).toMatchObject({ ok: false, code: 'pivot', message: 'PivotTable field name already exists: "n"' })
      expect(await s.setCellText('P', 'A4', '=1')).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('cannot enter a formula') })
      await expect(s.setCells([{ sheet: 'P', cell: 'A4', input: null }])).rejects.toThrow(/affect a PivotTable/)
      await expect(s.setCells([{ sheet: 'P', cell: 'B1', input: { value: 'b' } }])).rejects.toThrow('"b" is not an item of this field')
      // a new name is an edit of the definition, which the session makes with editPivotTable
      await expect(s.setCellText('P', 'A4', 'z')).rejects.toThrow(/editPivotTable/)
      await s.close()
    } finally {
      setPivotEntry(() => null)
    }
  })
})
