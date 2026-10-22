/**
 * Reproductions from the 21.10 review, reading: `findCells` in a cell's value reads a logical as TRUE or FALSE and an
 * error by its name, in any locale. In a synchronous session, in the same thread and in a thread, with closures and eval.
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
  })
})
