/**
 * The text a window's cell opens with for editing (`ViewportOptions.editText`), as Excel's formula bar shows it: a
 * formula with `=` in the book's locale (inside its table as Excel shows it there), a value as `entryText` writes it by
 * the format the cell shows (its own, its row's or its column's), text that would read as something else with an
 * apostrophe, by the session's clock for a date without a year; empty for an empty cell with a style and a cell an
 * array spilled into; the area's formula in every cell of an old-kind array formula's area. In a synchronous session, in
 * the same thread and in a thread, with closures and eval, and over a saved result (`viewportOf`).
 */
import { afterAll, describe, expect, it } from 'vitest'
import { entryText, viewportOf, type BookInput, type Execution, type ViewportCell } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

/** Formulas in ru-RU, in a table too, values under formats of their own and of their row and column, texts that read as values, styled empty cells, a spill. */
const book = (): BookInput => ({
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { formula: 'ROUND(SUM(B2:B3)*1.5,0)' }, A2: { value: 46096, numberFormat: 14 }, A3: { value: 45688.75, numberFormat: 22 }, A4: { value: '007' },
        A5: { value: '1,5' }, A6: { value: '1.5' }, A7: { value: 0.125, numberFormat: '0.0%' }, A8: { value: true }, A9: { value: '29.02' }, A10: { value: 1234.5 }, A11: { value: 'текст', style: 0 },
        B1: { value: 'Qty' }, B2: { value: 2 }, B3: { value: 3 }, C1: { value: 'Total' }, C2: { formula: '[@Qty]*2' }, C3: { formula: '[@Qty]*2' },
        D1: { formula: 'SEQUENCE(2)' }, E2: { value: 0.5 }, E4: { value: 45688 }, E5: { value: 45688 },
      },
      rows: { 4: { numberFormat: '0%' }, 12: { style: 0 } },
      columns: { E: { numberFormat: 'dd.mm.yyyy' } },
      tables: [{ name: 'T', range: 'B1:C3', columns: [{ name: 'Qty' }, { name: 'Total', formula: '[@Qty]*2' }] }],
    },
  ],
})

const texts = (cells: Record<string, ViewportCell>) => Object.fromEntries(Object.entries(cells).map(([a, c]) => [a, c.editText]))

describe.each(['closures', 'eval'] as const)('the text a cell opens with, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('a formula in the locale, a value by the format it shows, an apostrophe where text reads otherwise, nothing for an empty or spilled cell', async () => {
      const open = await opener(kind)
      const s = await open(book(), { clock: CLOCK, execution })
      const v = await s.getViewport('S', 'A1:E12', { editText: true })
      expect(texts(v.cells)).toMatchObject({
        A1: '=ОКРУГЛ(СУММ(B2:B3)*1,5;0)', A2: '15.03.2026', A3: '31.01.2025  18:00:00', A4: "'007", A5: "'1,5", A6: "'1.5", A7: '12,5%', A8: 'ИСТИНА',
        A9: '29.02', A10: '1234,5', A11: 'текст', C2: '=[@Qty]*2', D1: '=ПОСЛЕД(2)', D2: '', E2: '00.01.1900  12:00:00', E4: '4568800%', E5: '31.01.2025', A12: '', B12: '',
      })
      // a value under its row's percent and its column's date format, and every value as entryText writes it
      expect(v.cells.E4).toMatchObject({ numberFormat: '0%', numberFormatSource: 'row', editText: '4568800%' })
      for (const [a, c] of Object.entries(v.cells))
        if (c.formula === undefined && c.spillOf === undefined && c.value !== null) expect(c.editText, a).toBe(entryText(c.value, { locale: 'ru-RU', numberFormat: c.numberFormat, now: CLOCK.now }))
      expect(Object.values((await s.getViewport('S', 'A1:E12')).cells).some((c) => 'editText' in c)).toBe(false)
      // the year of the session's clock: 29.02 is a date in a leap year
      const leap = await open(book(), { clock: { now: '2028-03-01T00:00:00Z', seed: 1 }, execution })
      expect((await leap.getViewport('S', 'A9', { editText: true })).cells.A9.editText).toBe("'29.02")
      await leap.close()
      await s.close()
    })

    it("every cell of an old-kind array formula's area opens with its formula, as F2 opens it; a cell a dynamic array spilled into opens empty; a saved result's window alike", async () => {
      const b = book()
      Object.assign(b.sheets[0].cells, { G1: { formula: 'B2:B3*2', array: 'G1:G2' }, H1: { formula: 'SUM(B2:B3)', array: 'H1:I2' }, J1: { formula: 'SUM(B2:B3*1.5)', array: 'J1' } })
      const s = await (await opener(kind))(b, { clock: CLOCK, execution })
      const want = { D1: '=ПОСЛЕД(2)', D2: '', G1: '=B2:B3*2', G2: '=B2:B3*2', H1: '=СУММ(B2:B3)', I1: '=СУММ(B2:B3)', H2: '=СУММ(B2:B3)', I2: '=СУММ(B2:B3)', J1: '=СУММ(B2:B3*1,5)' }
      expect(texts((await s.getViewport('S', 'D1:J2', { editText: true })).cells)).toMatchObject(want)
      expect(texts(viewportOf(await s.getSnapshot(), 'S', 'D1:J2', { editText: true }).cells)).toMatchObject(want)
      await s.close()
    })

    it("a saved result's window: the same texts, by the result's clock", async () => {
      const s = await (await opener(kind))(book(), { clock: CLOCK, execution })
      const live = texts((await s.getViewport('S', 'A1:E12', { editText: true })).cells)
      expect(texts(viewportOf(await s.getSnapshot(), 'S', 'A1:E12', { editText: true }).cells)).toEqual(live)
      const result = await s.getResult()
      expect(viewportOf({ book: await s.getBook(), result: { ...result, clock: { now: '2028-03-01T00:00:00Z', seed: 1 } } }, 'S', 'A9', { editText: true }).cells.A9.editText).toBe("'29.02")
      await s.close()
    })
  })
})
