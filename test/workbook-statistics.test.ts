/**
 * `getWorkbookStatistics` in every kind of session, with closures and eval: a sheet's end and counts and the
 * workbook's, after edits and undo; a sheet that does not exist refused. Excel's own counts: `test/stats-excel.test.ts`.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type BookInput, type Execution } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({
  sheets: [
    { name: 'S', cells: { A1: { value: 1, note: { text: 'n' } }, B2: { formula: 'SEQUENCE(3)' }, D1: { formula: 'ROW(D1:D2)', array: 'D1:D2' }, F1: { value: 'x' } }, tables: [{ name: 'Data', range: 'F1:F3', columns: [{ name: 'x' }] }] },
    { name: 'H', hidden: true, cells: { A1: { value: 'x' }, C5: { note: { text: 'h' } } } },
  ],
})

const UNHELD = { pivotTables: null, charts: null, images: null, formControls: null, objects: null, comments: null }

describe.each(['closures', 'eval'] as Execution[])('%s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it("counts a sheet's cells, formulas, tables and notes, ends it at its last row and column of data or notes, and sums the workbook", async () => {
      const s = await (await opener(kind))(book(), { execution })
      // A1, B2:B4 spilled, D1:D2 a legacy array, the table's header F1
      expect(await s.getWorkbookStatistics('S')).toEqual({
        sheet: { end: 'F4', cells: 7, tables: 1, formulas: 5, notes: 1, ...UNHELD },
        workbook: { sheets: 2, cells: 8, tables: 1, formulas: 5, notes: 2, ...UNHELD },
      })
      expect((await s.getWorkbookStatistics('H')).sheet).toEqual({ end: 'C5', cells: 1, tables: 0, formulas: 0, notes: 1, ...UNHELD })
      await s.setCells([{ sheet: 'S', cell: 'B2', input: { formula: 'SEQUENCE(5)' } }])
      expect((await s.getWorkbookStatistics('S')).sheet).toMatchObject({ end: 'F6', cells: 9, formulas: 7 })
      await s.undo()
      expect((await s.getWorkbookStatistics('S')).sheet).toMatchObject({ end: 'F4', cells: 7, formulas: 5 })
      await expect(s.getWorkbookStatistics('Nope')).rejects.toThrow(InputError)
      await s.close()
    })
  })
})

it('a sheet without data or notes has no end, whatever its styles, lines and hyperlinks', async () => {
  const s = await (await opener(KINDS[0]))({ styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { H20: { style: 0 }, C3: { link: { url: 'https://example.com/' } } }, rows: { '30': { height: 30 } } }] }, {})
  expect((await s.getWorkbookStatistics('S')).sheet).toEqual({ end: null, cells: 0, tables: 0, formulas: 0, notes: 0, ...UNHELD })
  await s.close()
})
