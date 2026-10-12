/**
 * Links to other workbooks in every kind of session and both execution modes: an edit of a cell a formula with a link
 * depends on computes again as the full computation does (the hashes agree), a formula entered with a link marks that
 * book (`refreshError`) and its cells the stored values lack become `#REF!`, and a formula typed as Excel shows a link
 * is numbered by the file it names, one to a book the links lack refused.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type BookInput, type Clock } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

const clock: Clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

const BOOK: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2 },
        B1: { formula: "A1*'[1]Sheet1'!A1" },
        B2: { formula: "SUM('[1]Sheet1'!A:A)+A1" },
        B3: { formula: "'[1]Sheet1'!A9" },
        B4: { formula: "VLOOKUP(A1,'[1]Sheet1'!A1:B3,2,FALSE)" },
        B5: { formula: 'Ext+A1' },
        B6: { formula: "'[2]Data'!A9" },
      },
    },
  ],
  names: [{ name: 'Ext', ref: "'[1]Sheet1'!$A$3" }],
  externalBooks: [
    { target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1, A2: 2, A3: 3, B1: 'one', B2: 'two', B3: 'three' } }] },
    { target: '../data/Data.xlsx', sheets: [{ name: 'Data', cells: { A1: 7 } }] },
  ],
}

afterAll(closeCalcs)

describe.each(KINDS)('links in %s', (kind) => {
  describe.each(['closures', 'eval'] as const)('%s', (execution) => {
    it('an edit of a cell formulas with links read computes as the full computation', async () => {
      const s = await (await opener(kind))(BOOK, { clock, execution })
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      const r = await s.getResult()
      expect([1, 2, 3, 4, 5, 6].map((i) => r.sheets[0].cells[`B${i}`].value)).toEqual([3, 9, 0, 'three', 6, 0])
      expect(r.hash).toBe(compute(await s.getBook(), { clock, execution }).hash)
    })

    it('a formula entered with a link marks that book alone, its cells it lacks #REF! from then on', async () => {
      const s = await (await opener(kind))(BOOK, { clock, execution })
      await s.setCellText('S', 'C1', '=[Book2.xlsx]Sheet1!A2')
      const book = await s.getBook()
      expect(book.sheets[0].cells.C1.formula).toBe("'[1]Sheet1'!A2")
      expect(book.externalBooks!.map((b) => b.sheets.map((x) => x.refreshError ?? false))).toEqual([[true], [false]])
      const r = await s.getResult()
      expect(['B2', 'B3', 'B6', 'C1'].map((a) => r.sheets[0].cells[a].value)).toEqual([{ error: 'REF' }, { error: 'REF' }, 0, 2])
      expect(r.hash).toBe(compute(book, { clock, execution }).hash)
    })

    it('a formula typed with a book the links lack is refused, as a new link', async () => {
      const s = await (await opener(kind))(BOOK, { clock, execution })
      const r = await s.setCellText('S', 'C1', '=[Other.xlsx]Sheet1!A1')
      expect(r.ok).toBe(false)
      expect(!r.ok && r.code).toBe('external_ref')
      expect((await s.getBook()).sheets[0].cells.C1).toBeUndefined()
    })
  })
})
