/**
 * A number format as a code beside the format (`numberFormatCode`): on a cell of `getCells` and `getCell` with a format of
 * its own, and on a change whose format changed to one (`setCells`, `patchCells`, `setCellText`, an undo), a built-in
 * format by the code of the book's locale as the window gives it (14 is `dd.mm.yyyy` in ru-RU, `m/d/yyyy` in en-US), a
 * code as it is; none for a format removed or one the cell shows from its line, without the display too. In a synchronous
 * session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { builtinFormatCode, type BookInput, type Execution } from '../src/index.js'
import { CLOCK } from './async-books.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'S', cells: { A1: { value: 46096, numberFormat: 14 }, B1: { value: 0.5, numberFormat: '0.0%' }, C1: { formula: 'A1+1' }, D1: { value: 3 } }, columns: { D: { numberFormat: 4 } } }],
})

describe.each(['closures', 'eval'] as const)("a number format's code, execution %s", (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it("on a cell's own format, as the window gives it, in the book's locale", async () => {
      const open = await opener(kind)
      for (const [locale, date] of [['ru-RU', 'dd.mm.yyyy'], [undefined, 'm/d/yyyy']] as const) {
        const s = await open(book(locale), { clock: CLOCK, execution })
        const cells = await s.getCells('S', 'A1:D1')
        expect([cells.A1.numberFormatCode, cells.B1.numberFormatCode, cells.C1.numberFormatCode, cells.D1.numberFormatCode]).toEqual([date, '0.0%', undefined, undefined])
        expect(date).toBe(builtinFormatCode(14, locale))
        const v = await s.getViewport('S', 'A1:B1')
        expect([v.cells.A1.numberFormat, v.cells.B1.numberFormat]).toEqual([cells.A1.numberFormatCode, cells.B1.numberFormatCode])
        expect(await s.getCell('S', 'A1')).toEqual(cells.A1)
        await s.close()
      }
    })

    it('on a change whose format changed to one; none for a format removed; without the display too', async () => {
      const open = await opener(kind)
      for (const display of [true, false]) {
        const s = await open(book('ru-RU'), { clock: CLOCK, execution, display })
        const set = await s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 0.25, numberFormat: 10 } }, { sheet: 'S', cell: 'C1', input: { numberFormat: '0.000' } }])
        expect(set.changes.map((c) => [c.cell, c.numberFormat, c.numberFormatCode])).toEqual([['A2', 10, '0.00%'], ['C1', '0.000', '0.000']])
        const patched = await s.patchCells([{ sheet: 'S', cell: 'A1', set: { numberFormat: null } }])
        expect(patched.changes.map((c) => [c.cell, c.numberFormat, c.numberFormatCode])).toEqual([['A1', null, undefined]])
        const back = await s.undo()
        expect(back!.changes.map((c) => [c.cell, c.numberFormat, c.numberFormatCode])).toEqual([['A1', 14, 'dd.mm.yyyy']])
        // the date widens column A: A1, a date too, shows it instead of ###; its format did not change
        const typed = await s.setCellText('S', 'A3', '15.03.2026')
        expect(typed.ok && typed.changes.map((c) => [c.cell, c.numberFormat, c.numberFormatCode])).toEqual([['A3', 14, 'dd.mm.yyyy'], ...(display ? [['A1', undefined, undefined]] : [])])
        // a value typed into a cell that shows its column's format takes no format of its own, and no code
        const line = await s.setCellText('S', 'D2', '5')
        expect(line.ok && line.changes.map((c) => [c.cell, c.numberFormat, c.numberFormatCode])).toEqual([['D2', undefined, undefined]])
        await s.close()
      }
    })
  })
})
