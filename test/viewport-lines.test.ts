/**
 * A window by some rows and columns of its range (`ViewportOptions.rows`, `columns`), as a grid under a filter draws
 * it: only those lines, in their order, each with its number (as every window's lines have it); only their cells, each
 * as the whole window gives it; the merged areas and tables that meet them, a merged area over rows left out too;
 * `maxCells` counted by the cells given; the numbers checked. In a synchronous session, in the same thread and in a
 * thread, with closures and eval, and over a saved result (`viewportOf`).
 */
import { afterAll, describe, expect, it } from 'vitest'
import { createSession, InputError, viewportOf, type BookInput, type Execution, type Viewport, type ViewportOptions } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { randomBook } from './seeded-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { spillBook } from './session-spill-edits.js'

afterAll(closeCalcs)

/** Values in A1:F10 (`10 * row + column`), merged areas over two rows and in one, a table, sized and hidden lines, a rule. */
const grid = (): BookInput => {
  const cells: BookInput['sheets'][number]['cells'] = {}
  for (let r = 1; r <= 10; r++) for (let c = 1; c <= 6; c++) cells[`${'ABCDEF'[c - 1]}${r}`] = { value: 10 * r + c }
  cells.D7 = { value: 'Key' }
  cells.E7 = { value: 'Qty' }
  cells.E9 = { formula: 'B1*2' }
  return {
    styles: [{ font: { bold: true } }],
    sheets: [
      {
        name: 'S',
        cells,
        rows: { 4: { height: 30 }, 6: { hidden: true }, 9: { style: 0 } },
        columns: { C: { width: 20 }, F: { hidden: true } },
        merges: ['B2:C3', 'E5:F5', 'A8:A9'],
        tables: [{ name: 'T', range: 'D7:E9', columns: [{ name: 'Key' }, { name: 'Qty' }] }],
        conditionalFormats: [{ type: 'cellIs', ranges: ['A1:F10'], operator: 'greaterThan', formulas: ['50'], style: { fill: { color: '#FFC7CE' } } }],
      },
    ],
  }
}

/** The whole window's answer for the lines `rows` and `cols` (numbers from 1): what a window of those lines gives. */
function restricted(v: Viewport, rows: number[], cols: number[]) {
  const [rs, cs] = [new Set(rows), new Set(cols)]
  return {
    cells: Object.fromEntries(Object.entries(v.cells).filter(([a]) => rs.has(Number(a.replace(/^[A-Z]+/, ''))) && cs.has(a.replace(/\d+$/, '').split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)))),
    rows: v.rows.filter((r) => rs.has(r.row)),
    columns: v.columns.filter((c) => cs.has(c.column)),
  }
}

/** A window's cells and lines. */
const lined = (v: Viewport) => ({ cells: v.cells, rows: v.rows, columns: v.columns })

describe.each(['closures', 'eval'] as const)('a window by some rows and columns, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('only those lines in order, each numbered, their cells as the whole window has them, the merged areas and tables that meet them', async () => {
      const s = await (await opener(kind))(grid(), { clock: CLOCK, execution })
      const v = await s.getViewport('S', 'A1:F10', { rows: [1, 3, 5, 9], columns: [2, 5, 6] })
      expect(v.rows).toEqual([{ row: 1, height: 15 }, { row: 3, height: 15 }, { row: 5, height: 15 }, { row: 9, height: 15 }])
      expect(v.columns).toEqual([{ column: 2, width: 8.83203125 }, { column: 5, width: 8.83203125 }, { column: 6, width: 8.83203125, hidden: true }])
      expect(Object.keys(v.cells)).toEqual(['B1', 'E1', 'F1', 'B3', 'E3', 'F3', 'B5', 'E5', 'F5', 'B9', 'E9', 'F9'])
      // B2:C3 reaches a row left out; A8:A9 meets no column given
      expect(v.merges).toEqual(['B2:C3', 'E5:F5'])
      expect(v.tables).toEqual([{ name: 'T', range: 'D7:E9', headers: true, totals: false }])
      expect(v).toMatchObject({ range: 'A1:F10', truncated: false, usedRange: { rows: 10, columns: 6 }, frozen: { rows: 0, columns: 0 }, styles: { 0: { font: { bold: true } } } })
      expect(v.cells.E9).toEqual({ value: 24, text: '24', formula: 'B1*2', effectiveStyle: 0 })
      expect(v.cells.F5.appearance).toEqual({ style: { fill: { color: '#FFC7CE' } } })
      expect(lined(v)).toEqual(restricted(await s.getViewport('S', 'A1:F10'), [1, 3, 5, 9], [2, 5, 6]))
      // without the lines of the merged areas and the table
      expect(await s.getViewport('S', 'A1:F10', { rows: [1, 4], columns: [1, 4] })).toMatchObject({ merges: [], tables: [], rows: [{ row: 1 }, { row: 4, height: 30 }] })
      await s.close()
    })

    it('rows alone, columns alone, none, a whole column; every window numbers its lines', async () => {
      const s = await (await opener(kind))(grid(), { clock: CLOCK, execution })
      const rows = await s.getViewport('S', 'B2:C9', { rows: [2, 6, 7] })
      expect([rows.rows.map((r) => r.row), rows.columns.map((c) => c.column), Object.keys(rows.cells)]).toEqual([[2, 6, 7], [2, 3], ['B2', 'C2', 'B6', 'C6', 'B7', 'C7']])
      expect(rows.rows[1]).toEqual({ row: 6, height: 15, hidden: true })
      const cols = await s.getViewport('S', 'A9:F10', { columns: [3] })
      expect([cols.rows.map((r) => r.row), cols.columns, Object.keys(cols.cells)]).toEqual([[9, 10], [{ column: 3, width: 20 }], ['C9', 'C10']])
      const none = await s.getViewport('S', 'A1:F10', { rows: [] })
      expect(none).toMatchObject({ range: 'A1:F10', truncated: false, cells: {}, rows: [], merges: [], tables: [] })
      expect(none.columns).toHaveLength(6)
      const far = await s.getViewport('S', 'B:B', { rows: [3, 1_048_576] })
      expect([far.range, far.truncated, far.rows.map((r) => r.row), Object.keys(far.cells)]).toEqual(['B1:B1048576', false, [3, 1_048_576], ['B3']])
      expect((await s.getViewport('S', 'C4:D5')).rows).toEqual([{ row: 4, height: 30 }, { row: 5, height: 15 }])
      await s.close()
    })

    it('maxCells counts the cells given: the first rows asked that fit, or the first cells of the first one', async () => {
      const s = await (await opener(kind))(grid(), { clock: CLOCK, execution })
      const o: ViewportOptions = { rows: [1, 3, 5, 7, 9], columns: [2, 5, 6] }
      const two = await s.getViewport('S', 'A1:F10', { ...o, maxCells: 7 })
      expect([two.range, two.truncated, two.rows.map((r) => r.row), two.columns.length]).toEqual(['A1:F3', true, [1, 3], 3])
      const one = await s.getViewport('S', 'A1:F10', { ...o, maxCells: 2 })
      expect([one.range, one.truncated, one.rows.map((r) => r.row), one.columns.map((c) => c.column), Object.keys(one.cells)]).toEqual(['A1:E1', true, [1], [2, 5], ['B1', 'E1']])
      expect(await s.getViewport('S', 'A1:F10', { ...o, maxCells: 15 })).toMatchObject({ range: 'A1:F10', truncated: false })
      await s.close()
    })

    it('over other books: each window by lines has the cells the whole one has there', async () => {
      const open = await opener(kind)
      for (const book of [spillBook(), randomBook(5)]) {
        const s = await open(book, { clock: CLOCK, execution })
        for (const sheet of (await s.getBook()).sheets) {
          const [rows, columns] = [[1, 2, 5, 9, 12], [1, 3, 4, 10]]
          expect(lined(await s.getViewport(sheet.name, 'A1:J12', { rows, columns, diagnostics: true, formulaLocal: true })), sheet.name).toEqual(restricted(await s.getViewport(sheet.name, 'A1:J12', { diagnostics: true, formulaLocal: true }), rows, columns))
        }
        await s.close()
      }
    })
  })
})

describe('the rows and columns of a window', () => {
  it('are numbers of the range, ascending, each once', () => {
    const s = createSession(grid(), { clock: CLOCK })
    expect(() => s.getViewport('S', 'B2:D9', { rows: [1] })).toThrow('opts.rows[0]: row 1 is outside the range, rows 2 to 9')
    expect(() => s.getViewport('S', 'B2:D9', { columns: [2, 5] })).toThrow('opts.columns[1]: column 5 is outside the range, columns 2 to 4')
    expect(() => s.getViewport('S', 'B2:D9', { rows: [3, 3] })).toThrow('opts.rows[1]: expected numbers in ascending order, each once: 3 comes after 3')
    expect(() => s.getViewport('S', 'B2:D9', { rows: [5, 3] })).toThrow('opts.rows[1]')
    expect(() => s.getViewport('S', 'B2:D9', { rows: [2.5] })).toThrow('opts.rows[0]: expected an integer from 1 to 1048576, got number 2.5')
    expect(() => s.getViewport('S', 'B2:D9', { columns: 'B' as never })).toThrow(InputError)
    expect(() => s.getViewport('S', 'B2:D9', { rows: [2], columns: [] })).not.toThrow()
  })

  it('a saved window by lines is the session\'s', () => {
    const s = createSession(grid(), { clock: CLOCK })
    const o: ViewportOptions = { rows: [2, 3, 8], columns: [1, 2, 6], maxCells: 6 }
    const { revision: _r, ...live } = s.getViewport('S', 'A1:F10', o)
    for (const c of Object.values(live.cells)) delete c.appearance
    expect(viewportOf(s.getSnapshot(), 'S', 'A1:F10', o)).toEqual(live)
    expect(live).toMatchObject({ range: 'A1:F3', truncated: true, merges: ['B2:C3'] })
  })
})
