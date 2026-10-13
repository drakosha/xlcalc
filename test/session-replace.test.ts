/**
 * `replaceInCells`, Excel's Replace All in a session: the query's matches as Find reads them (`?` one UTF-16 unit, `*`
 * as little as it can but the last), each cell found entered again as its own text edited (a percent without the
 * editor's `%`, a formula in the book's locale, a text that reads as a number kept a text), an array of the old kind
 * whole, a formula that no longer reads left as it was; hidden cells as Find skips them, the rows asked; one revision and
 * one step of the undo, none when nothing is found; the options checked. In a synchronous session, in the same thread
 * and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { BookInput, Execution } from '../src/index.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-13T12:00:00Z', seed: 1 }

const book: BookInput = {
  locale: 'ru-RU',
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'a😀b' }, A2: { value: 'a1b2b' }, A3: { value: '007' }, A4: { value: 0.05, numberFormat: '0%' }, A5: { formula: 'SUM(D1:D2)' }, A6: { value: 'hid a' },
        D1: { value: 1 }, D2: { value: 2 },
        H1: { formula: 'D1:D2*5', array: 'H1:H2' },
      },
      rows: { 6: { hidden: true } },
    },
    {
      name: 'F',
      cells: { A1: { value: 'x1' }, A2: { value: 'x2' }, A3: { value: 'x3' } },
      rows: { 2: { hidden: true } },
      filtered: true,
    },
  ],
}

const S = [{ sheet: 'S', range: 'A1:H9' }]
const W = { wildcards: true } as const

describe.each(['closures', 'eval'] as const)('replaceInCells, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('replaces as Excel replaces, in one step of the undo', async () => {
      const s = await (await opener(kind))(book, { clock, execution })
      const cell = async (a: string, sheet = 'S') => (await s.getBook()).sheets.find((x) => x.name === sheet)!.cells[a]
      // `?` is one UTF-16 unit; one revision, one step, undone at once
      const first = await s.replaceInCells(S, 'a??b', 'y', W)
      expect(first).toMatchObject({ revision: 1, replaced: 1, changes: [{ cell: 'A1', value: 'y' }] })
      expect(s.history.undo).toBe(1)
      await s.undo()
      expect(await cell('A1')).toEqual({ value: 'a😀b' })
      // `*` takes as little as it can
      await s.replaceInCells(S, 'a*b', 'x', W)
      expect([await cell('A1'), await cell('A2')]).toEqual([{ value: 'x' }, { value: 'x2b' }])
      // a percent as its own text edited, without the editor's `%`; a text that reads as a number stays a text
      await s.replaceInCells(S, '5', '6', W)
      expect(await cell('A4')).toEqual({ value: 0.06, numberFormat: '0%' })
      await s.replaceInCells(S, '%', '', W)
      expect(await cell('A4')).toEqual({ value: 6, numberFormat: '0%' })
      await s.replaceInCells(S, '7', '8', W)
      expect(await cell('A3')).toEqual({ value: '008' })
      // a formula in the book's locale; an array of the old kind whole from a cell of its area
      await s.replaceInCells(S, 'СУММ', 'МАКС', W)
      expect(await cell('A5')).toEqual({ formula: 'MAX(D1:D2)' })
      await s.replaceInCells([{ sheet: 'S', range: 'H2' }], '6', '7', W)
      expect(await cell('H1')).toEqual({ formula: 'D1:D2*7', array: 'H1:H2' })
      // a formula that no longer reads is left as it was, the other cells replaced
      const put = await s.replaceInCells(S, 'x', '(', W)
      expect([put.replaced, 'refused' in put]).toEqual([2, false])
      const refused = await s.replaceInCells(S, '(', '', W)
      expect(refused.refused).toMatchObject([{ sheet: 'S', cell: 'A5', code: 'syntax' }])
      expect([await cell('A1'), await cell('A2'), await cell('A5')]).toEqual([undefined, { value: '2b' }, { formula: 'MAX(D1:D2)' }])
      // nothing found: no revision, no step
      const at = s.revision
      expect(await s.replaceInCells(S, 'SUM', 'MAX', W)).toEqual({ revision: at, changes: [], replaced: 0 })
      expect(s.revision).toBe(at)
    })

    it('skips hidden cells as Find does, takes the rows asked, checks its options', async () => {
      const s = await (await opener(kind))(book, { clock, execution })
      const cells = async (sheet: string) => (await s.getBook()).sheets.find((x) => x.name === sheet)!.cells
      // a hidden row on a sheet without a filter is replaced; on a filtered sheet hidden cells are skipped
      await s.replaceInCells(S, 'hid', 'shown', { ...W, skipHidden: true })
      expect((await cells('S')).A6).toEqual({ value: 'shown a' })
      await s.replaceInCells([{ sheet: 'F', range: 'A1:A3' }], 'x', 'y', { ...W, skipHidden: true })
      expect(Object.values(await cells('F')).map((c) => c.value)).toEqual(['y1', 'x2', 'y3'])
      await s.replaceInCells([{ sheet: 'F', range: 'A1:A3' }], 'y', 'z', { rows: [3] })
      expect(Object.values(await cells('F')).map((c) => c.value)).toEqual(['y1', 'x2', 'z3'])
      await expect(s.replaceInCells(S, 'a', 'b', { in: 'edit' } as object)).rejects.toThrow('opts')
      await expect(s.replaceInCells(S, 'a', 5 as unknown as string)).rejects.toThrow('replacement')
      await expect(s.replaceInCells(S, 'a', 'b', { byColumns: true } as object)).rejects.toThrow('byColumns')
    })
  })
})
