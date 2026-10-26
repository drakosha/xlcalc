/**
 * `getAutoSum` in every kind of session: the range Excel's AutoSum offers and its formula in the workbook's locale
 * too. Excel's own answers are the cases of `test/excel/ui/press-held.json` and `press-more-held.json`
 * (`test/ui-excel.test.ts`); these cover the calls.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type BookInput } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import './locales.js'

const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'Amount' }, A2: { value: 1 }, A3: { value: 2 }, A4: { formula: 'SUM(A2:A3)' }, A5: { value: 4 },
        C2: { value: 1 }, D2: { value: 2 }, F1: { value: 45688, numberFormat: 14 },
      },
    },
  ],
})

afterAll(closeCalcs)

describe.each(KINDS)('getAutoSum, %s', (kind) => {
  const open = openerFor(kind)
  it('the numbers above, past a blank, up to a subtotal; to the left; subtotals alone; none for a date', async () => {
    const s = await open(book(), {})
    expect(await s.getAutoSum('S', 'A7')).toEqual({ range: 'A5:A6', formula: 'SUM(A5:A6)', formulaLocal: 'SUM(A5:A6)' })
    expect(await s.getAutoSum('S', 'A5')).toEqual({ range: 'A4', formula: 'SUM(A4)', formulaLocal: 'SUM(A4)' })
    expect((await s.getAutoSum('S', 'E2')).range).toBe('C2:D2')
    // left of G1, a date and text: no numbers
    expect((await s.getAutoSum('S', 'G1')).formula).toBe('SUM()')
    expect((await s.getAutoSum('S', 'F2')).range).toBe('C2:E2')
    await expect(s.getAutoSum('S', 'nope')).rejects.toThrow(InputError)
    s.close()
  })
  it("the formula in the workbook's locale", async () => {
    const s = await open(book('ru-RU'), {})
    expect(await s.getAutoSum('S', 'A7')).toEqual({ range: 'A5:A6', formula: 'SUM(A5:A6)', formulaLocal: 'СУММ(A5:A6)' })
    s.close()
  })
})
