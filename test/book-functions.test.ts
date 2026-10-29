/**
 * CELL, INFO, SHEET, SHEETS, FORMULATEXT and ISFORMULA over what the book holds beside values (`src/book-facts.ts`): the
 * file option in every kind of session and in both execution modes, CELL computed again when a look, a width or the
 * formula it reads changes, the active cell CELL reads without a reference. Excel's answers are in the sets `indirect`
 * and `indirect-held`.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type AsyncSession, type BookInput, type Execution } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-29T12:00:00Z', seed: 1 }
const file = { name: 'Budget.xlsx', path: '/Users/me/' }
afterAll(closeCalcs)

const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: -2.5 }, A2: { formula: 'A1*2' },
        B1: { formula: 'CELL("filename",A1)' }, B2: { formula: 'CELL("address",T!B3)' }, B3: { formula: 'INFO("directory")' },
        B4: { formula: 'CELL("format",A1)' }, B5: { formula: 'INDEX(CELL("width",A1),1,1)' }, B6: { formula: 'FORMULATEXT(A2)' }, B7: { formula: 'ISFORMULA(A2)' },
        B8: { formula: 'SHEET(T!A1)&"/"&SHEETS()' }, B9: { formula: 'CELL("address")' },
      },
    },
    { name: 'T', cells: {}, view: { activeCell: 'C4' } },
  ],
  activeSheet: 'T',
})

describe.each(['closures', 'eval'] as Execution[])('the book functions, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    const at = async (s: AsyncSession, refs = 'B1:B9') => {
      expect(await s.getHash()).toBe(compute(await s.getBook(), { clock, execution, file }).hash)
      return Object.values(await s.getCells('S', refs)).map((c) => c.value)
    }

    it('tell the file, the looks, the formulas and the sheets, and follow edits of them', async () => {
      const s = await open(book(), { clock, execution, file })
      try {
        expect(await at(s)).toEqual(['/Users/me/[Budget.xlsx]S', '[Budget.xlsx]T!$B$3', '/Users/me/', 'G', 8, '=A1*2', true, '2/2', '[Budget.xlsx]T!$C$4'])
        await s.patchCells([{ sheet: 'S', cell: 'A1', set: { numberFormat: '$#,##0.00_);[Red]($#,##0.00)' } }])
        expect(await at(s, 'B4')).toEqual(['C2-'])
        await s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 20.7109375 } }])
        expect(await at(s, 'B5')).toEqual([20])
        await s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: 'A1+1' } }])
        expect(await at(s, 'B6:B7')).toEqual(['=A1+1', true])
        await s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 3 } }])
        expect(await at(s, 'B6:B7')).toEqual([{ error: 'NA' }, false])
        await s.undo()
        await s.undo()
        expect(await at(s, 'B6')).toEqual(['=A1*2'])
      } finally {
        await s.close()
      }
    })
  })
})

describe('without the file option', () => {
  it('is a workbook not saved yet: no file name, Book1 in an address of another sheet, no folder', () => {
    const r = compute(book(), { clock })
    expect(['B1', 'B2', 'B3'].map((a) => r.sheets[0].cells[a].value)).toEqual(['', '[Book1]T!$B$3', ''])
  })

  it('INFO answers constants and the book calculation', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'INFO("system")&INFO("release")&INFO("recalc")' } } }], calculation: 'manual' }, { clock })
    expect(r.sheets[0].cells.A1.value).toBe('pcdos16.0Manual')
  })
})
