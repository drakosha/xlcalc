/**
 * `getSnapshot`: a session's book and its result in one call, in a synchronous session, in the same thread and in a
 * thread, with closures and eval: what `getBook()` and `getResult()` give, with no edit between the two (a call made before
 * it is in it, one made after it is not); written to xlsx as it is and read back with the session's values, and Excel's
 * values read from a file written again from the snapshot of a session over that file. Text typed into a cell of the
 * session's window keeps the cell's note and link (`parseEntry` with the cell's view).
 */
import { afterAll, describe, expect, it } from 'vitest'
import { parseEntry, type BookInput, type Execution } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { storedOf } from './xlsx-fixtures.js'

afterAll(closeCalcs)

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2, note: { author: 'Ann', text: 'From the CRM' }, link: { url: 'https://example.com/' }, style: 0 },
        A2: { value: 3 },
        B1: { formula: 'SEQUENCE(A1)*A2' },
        C1: { formula: 'SUM(B1#)' },
      },
    },
  ],
})

describe.each(KINDS)('getSnapshot, %s', (kind) => {
  describe.each(['closures', 'eval'] as Execution[])('%s', (execution) => {
    it('gives the book and its result, as book() and result() give them, the calls before it in it and none after', async () => {
      const open = await opener(kind)
      const s = await open(book(), { clock: CLOCK, execution })
      const edit = s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      const snap = s.getSnapshot()
      const later = s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 10 } }])
      const { book: b, result } = await snap
      await Promise.all([edit, later])
      expect(result.sheets[0].cells.C1.value).toBe(18)
      expect(result.sheets[0].cells.B1.spill).toBe('B1:B3')
      // from a thread the book is a copy; elsewhere the session's own objects, as book() gives them, which later calls change
      if (kind === 'a thread') expect([b.sheets[0].cells.A1.value, b.sheets[0].cells.A2]).toEqual([3, { value: 3 }])
      else expect(b).toBe(await s.getBook())
      const now = await s.getSnapshot()
      expect(now.result).toEqual({ ...(await s.getResult()), ms: now.result.ms })
      expect(now.book).toEqual(await s.getBook())
      expect(now.result.sheets[0].cells.C1.value).toBe(60)
      await s.close()
    })

    it('writes to xlsx as it is, and the file reads back with the session values', async () => {
      const open = await opener(kind)
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 4, style: 0, note: { text: 'n' } } }])
      const snap = await s.getSnapshot()
      const { data, issues } = await writeXlsx(snap)
      expect(issues).toEqual([])
      const read = await readXlsx(data)
      expect(read.book).toEqual(snap.book)
      expect(read.cached).toEqual(storedOf(snap.result))
      // a session over the file read gives a snapshot that writes the file again
      const again = await open(read.book, { clock: CLOCK, execution })
      expect((await writeXlsx(await again.getSnapshot())).data).toEqual(data)
      await Promise.all([s.close(), again.close()])
    })

    it("text typed into a cell of its window keeps the cell's note, link and style", async () => {
      const open = await opener(kind)
      const s = await open(book(), { clock: CLOCK, execution })
      const view = (await s.getCells('S', 'A1')).A1
      const typed = parseEntry('5', { cell: view })
      expect(typed).toEqual({ ok: true, input: { value: 5, style: 0, note: { author: 'Ann', text: 'From the CRM' }, link: { url: 'https://example.com/' } } })
      if (typed.ok) await s.setCells([{ sheet: 'S', cell: 'A1', input: typed.input }])
      expect((await s.getCells('S', 'A1')).A1).toMatchObject({ value: 5, note: { author: 'Ann', text: 'From the CRM' }, link: { url: 'https://example.com/' } })
      await s.close()
    })
  })
})
