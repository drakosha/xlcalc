/**
 * Text typed into a session's cell (`setCellText`, `previewCellText`): read in the book's locale, date system and clock by
 * the format the cell shows, a format its line gives it not written into it, its look kept; the column widened in the
 * same step as Excel widens it, or the width only told; a value its rule of data validation does not take written with
 * the rule; a formula that does not read, part of an array and a formula in a table's header refused with a message in
 * the book's locale; a header typed over renamed. In a synchronous session, in the same thread and in a thread, with
 * closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type BookInput, type Execution } from '../src/index.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1, style: 0, meta: 'm', note: { text: 'n' } },
        B1: { formula: 'A1*2' },
        C1: { formula: 'A1:A2*10', array: 'C1:C2' },
        E1: { value: 'Qty' },
        E2: { value: 1 },
      },
      columns: { D: { format: '0%' }, F: { width: 4 } },
      rows: { 5: { format: 'dd.mm.yyyy' } },
      tables: [{ name: 'T', range: 'E1:E2', columns: [{ name: 'Qty' }] }],
      validations: [{ range: 'G1:G9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning', error: 'Positive only' }],
    },
  ],
})

describe.each(['closures', 'eval'] as const)('text typed into a cell, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it("read in the book's locale and clock by the format the cell shows; the cell keeps its look; a line's format stays the line's", async () => {
      const s = await (await opener(kind))(book('ru-RU'), { clock, execution })
      // a grouped number takes Excel's grouped format, the cell keeps its style, meta and note
      const r = await s.setCellText('S', 'A1', '1 234,5')
      expect(r).toEqual({
        ok: true,
        revision: 1,
        changes: [
          { sheet: 'S', cell: 'A1', value: 1234.5, text: '1 234,50', numberFormat: 4 },
          { sheet: 'S', cell: 'C1', value: 12345, text: '12345', spill: 'C1:C2' },
          { sheet: 'S', cell: 'B1', value: 2469, text: '2469' },
        ],
      })
      const cells = () => s.book().then((b) => b.sheets[0].cells)
      expect((await cells()).A1).toEqual({ value: 1234.5, format: 4, style: 0, meta: 'm', note: { text: 'n' } })
      // a number into a percent column is percent, by the column's format, which the cell does not take
      await s.setCellText('S', 'D2', '5')
      expect((await cells()).D2).toEqual({ value: 0.05 })
      // a date without a year takes the session's clock's; a row of dates keeps its format; the column widens for it
      expect(await s.setCellText('S', 'A5', '15.3')).toMatchObject({ changes: [{ cell: 'A5', value: 46096, text: '15.03.2026' }], width: 10.1640625 })
      expect((await cells()).A5).toEqual({ value: 46096 })
      // a formula in the locale's words is stored in English
      expect((await s.setCellText('S', 'A6', '=СУММ(A1;1,5)')).ok).toBe(true)
      expect((await cells()).A6).toEqual({ formula: 'SUM(A1,1.5)' })
      // empty text clears the value and keeps the rest
      await s.setCellText('S', 'A1', '')
      expect((await cells()).A1).toEqual({ format: 4, style: 0, meta: 'm', note: { text: 'n' } })
      expect(await s.hash()).toBe(compute(await s.book(), { clock, execution }).hash)
      await s.close()
    })

    it('a number too wide for its column widens it in the same step, or only tells the width; a column set by hand stays', async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      expect(await s.previewCellText('S', 'B9', '1234567890')).toEqual({ ok: true, input: { value: 1234567890 }, value: 1234567890, width: 11.1640625 })
      const r = await s.setCellText('S', 'B9', '1234567890')
      expect(r).toMatchObject({ ok: true, revision: 1, width: 11.1640625, layout: ['S'] })
      expect((await s.layout()).sheets[0].columns!.B).toEqual({ width: 11.1640625, bestFit: true })
      const told = await s.setCellText('S', 'H1', '1234567890', { widen: false })
      expect(told).toMatchObject({ ok: true, width: 11.1640625 })
      expect((await s.layout()).sheets[0].columns!.H).toBeUndefined()
      expect(await s.setCellText('S', 'F1', '1234567890')).not.toHaveProperty('width')
      // the width goes back with the value, in one step
      expect((await s.undo())!.call).toBe('setCellText')
      await s.undo()
      expect((await s.undo())!.layout).toEqual(['S'])
      expect((await s.layout()).sheets[0].columns!.B).toBeUndefined()
      await s.close()
    })

    it("a value the cell's rule does not take is written, with the rule's alert; the preview tells it with nothing written", async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      const rule = book().sheets[0].validations![0]
      expect(await s.previewCellText('S', 'G2', '-3')).toEqual({ ok: true, input: { value: -3 }, value: -3, validation: { valid: false, rule, severity: 'warning', message: 'Positive only' } })
      expect((await s.book()).sheets[0].cells.G2).toBeUndefined()
      expect(s.revision).toBe(0)
      const r = await s.setCellText('S', 'G2', '-3')
      expect(r.ok && r.validation).toEqual({ valid: false, rule, severity: 'warning', message: 'Positive only' })
      expect((await s.book()).sheets[0].cells.G2).toEqual({ value: -3 })
      expect(await s.setCellText('S', 'G3', '4')).not.toHaveProperty('validation')
      expect(await s.validateEntry('S', 'G3', { value: 0 })).toEqual({ ok: false, code: 'validation', message: 'Positive only', rule, severity: 'warning' })
      await s.close()
    })

    it("refused: a formula that does not read, part of an array, a formula in a table's header; a header typed over renames its column", async () => {
      const s = await (await opener(kind))(book('ru-RU'), { clock, execution })
      expect(await s.setCellText('S', 'A3', '=СУММ(A1;1.5)')).toEqual({ ok: false, code: 'syntax', message: 'Здесь не может стоять «.»', at: 10 })
      expect(await s.setCellText('S', 'C2', '7')).toEqual({ ok: false, code: 'array', message: 'Нельзя изменить часть массива S!C1 (C1:C2)' })
      expect(await s.setCellText('S', 'E1', '=1+1')).toEqual({ ok: false, code: 'header', message: 'Ячейка S!E1 в строке заголовков таблицы «T»: в ней имя столбца, а не формула' })
      expect([s.revision, s.history.undo]).toEqual([0, 0])
      const r = await s.setCellText('S', 'E1', 'Кол-во')
      expect(r).toMatchObject({ ok: true, revision: 1, layout: ['S'] })
      expect((await s.book()).sheets[0].tables![0].columns[0].name).toBe('Кол-во')
      await s.close()
    })
  })
})
