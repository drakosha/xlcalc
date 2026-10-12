/**
 * `findCells` by Excel's Find in a session: the query's wildcards (`?` one UTF-16 unit), the whole cell, Match case,
 * Look in Formulas (`in: 'edit'`, the book's locale, a text without its apostrophe) and Values (`in: 'text'`), the
 * cells by rows or by columns with pages and a count, hidden cells as Find skips them (Values always, every search on a
 * filtered sheet, where cells come by blocks), and the options checked. In a synchronous session, in the same thread and
 * in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { BookInput, Execution, FindCellsOptions } from '../src/index.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

const book: BookInput = {
  locale: 'ru-RU',
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'a*b' }, B1: { value: 'axb' }, C1: { value: 'a😀b' },
        A2: { formula: 'SUM(D1:D2)' }, B2: { value: 0.05, numberFormat: '0%' }, C2: { value: '007' },
        D1: { value: 1 }, D2: { value: 2 },
        A3: { value: 'Straße' }, B3: { value: 'ΟΣ' }, C3: { value: 'hid' },
        A4: { value: 'hid' },
      },
      rows: { 4: { hidden: true } },
    },
    {
      name: 'F',
      cells: { A1: { value: 'x1' }, B1: { value: 'x2' }, C1: { value: 'x3' }, A2: { value: 'x4' }, B2: { value: 'x5' }, C2: { value: 'x6' }, A3: { value: 'x7' }, C3: { value: 'x8' } },
      rows: { 2: { hidden: true } },
      columns: { B: { hidden: true } },
      filtered: true,
    },
  ],
}

const S = [{ sheet: 'S', range: 'A1:D9' }]
const F = [{ sheet: 'F', range: 'A1:XFD1048576' }]

describe.each(['closures', 'eval'] as const)('findCells by Excel Find, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('reads wildcards, the whole cell and the case as Find; looks in formulas and in values', async () => {
      const s = await (await opener(kind))(book, { clock, execution })
      const cells = async (query: string, o: FindCellsOptions, ranges = S) => (await s.findCells(ranges, query, { ...o, fields: [] })).cells.map((c) => c.cell)
      const wild = { wildcards: true, in: 'edit' } as const
      // `?` is one UTF-16 unit: a character outside the BMP takes two; `~` takes the next character as it is
      expect(await cells('a?b', { ...wild, wholeCell: true })).toEqual(['A1', 'B1'])
      expect(await cells('a??b', { ...wild, wholeCell: true })).toEqual(['C1'])
      expect(await cells('a~*b', { ...wild, wholeCell: true })).toEqual(['A1'])
      // without wildcards `*` is itself
      expect(await cells('a*b', { in: 'edit' })).toEqual(['A1'])
      // any case: `ß` and `ẞ` are one letter, `σ`, `ς` and `Σ` too; Match case keeps them apart
      expect(await cells('STRAẞE', wild)).toEqual(['A3'])
      expect(await cells('ος', wild)).toEqual(['B3'])
      expect(await cells('ος', { ...wild, caseSensitive: true })).toEqual([])
      // Look in Formulas: the formula in the book's locale, a percent as typed, a text without its apostrophe
      expect(await cells('СУММ', wild)).toEqual(['A2'])
      expect(await cells('5%', wild)).toEqual(['B2'])
      expect(await cells("'", wild)).toEqual([])
      expect(await cells('0,05', wild)).toEqual([])
      // Look in Values: what cells show; the stored formula and the value as before
      expect(await cells('3', { wildcards: true, in: 'text', wholeCell: true })).toEqual(['A2'])
      expect(await cells('SUM(', { in: 'formula' })).toEqual(['A2'])
      // an empty query, and `~` alone, find nothing
      expect(await cells('', wild)).toEqual([])
      expect(await cells('~', wild)).toEqual([])
    })

    it('gives cells by rows or by columns in pages with a count; hidden cells as Find skips them', async () => {
      const s = await (await opener(kind))(book, { clock, execution })
      const first = await s.findCells(S, 'a', { wildcards: true, byColumns: true, maxCells: 2, count: true, fields: [] })
      expect(first).toMatchObject({ truncated: true, total: 4, cells: [{ cell: 'A1' }, { cell: 'A3' }] })
      const next = await s.findCells(S, 'a', { wildcards: true, byColumns: true, maxCells: 2, count: true, fields: [], cursor: first.cursor })
      expect(next.cells.map((c) => c.cell)).toEqual(['B1', 'C1'])
      expect(next.truncated).toBe(false)
      const hid = async (o: FindCellsOptions, ranges = S, query = 'hid') => (await s.findCells(ranges, query, { wildcards: true, skipHidden: true, fields: [], ...o })).cells.map((c) => c.cell)
      // a hidden row: Look in Formulas finds its cells, Values skips them
      expect(await hid({ in: 'edit' })).toEqual(['C3', 'A4'])
      expect(await hid({ in: 'text' })).toEqual(['C3'])
      // a filtered sheet: both skip hidden cells and take the rest by the blocks hidden lines leave, rows of blocks first
      expect(await hid({ in: 'edit' }, F, 'x')).toEqual(['A1', 'C1', 'A3', 'C3'])
      expect(await hid({ in: 'edit', byColumns: true }, F, 'x')).toEqual(['A1', 'C1', 'A3', 'C3'])
      expect(await hid({ in: 'text', skipHidden: false }, F, 'x')).toEqual(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'C3'])
    })

    it('checks its options', async () => {
      const s = await (await opener(kind))(book, { clock, execution })
      await expect(s.findCells(S, 'a', { in: 'formulas' as 'edit' })).rejects.toThrow('opts.in')
      await expect(s.findCells(S, 'a', { wildcards: 'yes' as unknown as boolean })).rejects.toThrow('opts.wildcards')
      await expect(s.findCells(S, 'a', { byRows: true } as FindCellsOptions)).rejects.toThrow('byRows')
    })
  })
})
