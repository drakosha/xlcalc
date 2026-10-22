/**
 * Reproductions from the 21.10 review, reading: `findCells` in a cell's value reads a logical as TRUE or FALSE and an
 * error by its name, in any locale. In a session opened with `display: false`, `readCells` and `findCells` return `text`
 * when their fields ask for it, as documented, and the cell views leave it out. In a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { BookInput } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'S', cells: { A1: { value: true }, A2: { formula: '1/0' }, A3: { value: 1.5 }, A4: { value: 'true' }, A5: { formula: 'FALSE()' } } }],
})

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)

    it.each([undefined, 'ru-RU'])("findCells in: 'value' reads TRUE, FALSE and #DIV/0! (locale %s)", async (locale) => {
      const s = await open(book(locale), { clock: CLOCK, execution })
      const found = async (query: string, caseSensitive = true) =>
        (await s.findCells([{ sheet: 'S', range: 'A1:A5' }], query, { in: 'value', caseSensitive, fields: ['value'] })).cells.map((x) => x.cell)
      expect(await found('TRUE')).toEqual(['A1'])
      expect(await found('true')).toEqual(['A4'])
      expect(await found('true', false)).toEqual(['A1', 'A4'])
      expect(await found('FALSE')).toEqual(['A5'])
      expect(await found('#DIV/0!')).toEqual(['A2'])
      // the error's internal code is not its text
      expect(await found('error', false)).toEqual([])
      expect(await found('1.5')).toEqual(['A3'])
      await s.close()
    })

    it('display: false: readCells and findCells give the text their fields ask for, getCell and getCells none', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 1.5, numberFormat: '0.00' } } }] }, { clock: CLOCK, execution, display: false })
      const at = [{ sheet: 'S', range: 'A1' }]
      expect((await s.readCells(at)).cells).toEqual([{ sheet: 'S', cell: 'A1', value: 1.5, text: '1.50' }])
      expect((await s.findCells(at, '1.50')).cells).toEqual([{ sheet: 'S', cell: 'A1', value: 1.5, text: '1.50' }])
      expect((await s.readCells(at, { fields: ['value', 'formula'] })).cells).toEqual([{ sheet: 'S', cell: 'A1', value: 1.5 }])
      expect(await s.getCell('S', 'A1')).not.toHaveProperty('text')
      expect((await s.getCells('S', 'A1')).A1).not.toHaveProperty('text')
      await s.close()
    })
  })
})
