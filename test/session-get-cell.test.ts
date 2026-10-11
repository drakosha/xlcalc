/**
 * `Session.getCell(sheet, cell)`: one cell as `getCells(sheet, cell)` gives it, every cell of a few rectangles of books
 * with values, formulas, spills, an old-kind array formula, styles, notes, links and conditional formats (an empty cell
 * with an appearance too); undefined for a cell `getCells` leaves out; a range or an unknown sheet an `InputError`. In a
 * synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type BookInput, type Execution } from '../src/index.js'
import { formatAddress } from '../src/model.js'
import { CLOCK } from './async-books.js'
import { randomBook } from './seeded-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { spillBook } from './session-spill-edits.js'

afterAll(closeCalcs)

const looks = (): BookInput => ({
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2, style: 0, note: { text: 'n' }, numberFormat: 14 },
        A2: { formula: 'SEQUENCE(3)' },
        B1: { formula: 'A1:A2*10', array: 'B1:B2' },
        C3: { value: 'x', link: { url: 'https://example.com/' }, meta: { id: 1 } },
      },
      conditionalFormats: [{ type: 'containsBlanks', ranges: ['D1:D3'], style: { fill: { color: '#FFC7CE' } } }],
    },
  ],
})

describe.each(['closures', 'eval'] as const)('one cell, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('is the cell getCells gives for it, or undefined', async () => {
      const open = await opener(kind)
      for (const book of [looks(), spillBook(), randomBook(3)]) {
        const s = await open(book, { clock: CLOCK, execution })
        for (const sheet of book.sheets)
          for (let r = 0; r < 8; r++)
            for (let c = 0; c < 6; c++) {
              const a = formatAddress(r, c)
              expect(await s.getCell(sheet.name, a), `${sheet.name}!${a}`).toEqual((await s.getCells(sheet.name, a))[a])
            }
        await s.close()
      }
      const s = await open(looks(), { clock: CLOCK, execution })
      expect(await s.getCell('S', 'A4')).toEqual({ value: 3, text: '3', spillOf: 'A2' })
      expect(await s.getCell('S', 'D2')).toEqual({ value: null, text: '', appearance: { style: { fill: { color: '#FFC7CE' } } } })
      expect(await s.getCell('S', 'E9')).toBeUndefined()
      await expect(s.getCell('S', 'A1:B2')).rejects.toThrow(new InputError('cell', 'invalid cell address "A1:B2": expected A1 style such as "B12"'))
      await expect(s.getCell('Nope', 'A1')).rejects.toThrow('sheet: unknown sheet "Nope"')
      await s.close()
    })
  })
})
