/**
 * `patchCells` with a style by its fields over a range: the fields go over the style each cell shows (its own, else its
 * row's, else its column's), null removes a field, equal styles share one entry of the book's table; a rectangle's empty
 * cells get a cell of their own; whole columns and rows are formatted as Excel formats them, the cells of theirs with a
 * look of their own and the cells where they cross a formatted line of the other direction merged; the number format by
 * `numberFormat` or its old name `format`; whole lines take no meta, notes or links; a rectangle past the bound is
 * refused. In a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, type BookInput, type Execution } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { cellsOf, expandSheet } from './looks-book.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }, { fill: { color: '#C6EFCE' } }, { font: { italic: true, color: '#C00000' } }],
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 1, style: 2 }, B2: { value: 2 }, C3: { value: 3, numberFormat: '0%' }, D5: { value: 4 } },
      columns: { C: { style: 1, numberFormat: '0.00' } },
      rows: { 5: { style: 0 } },
    },
  ],
})

describe.each(['closures', 'eval'] as const)('styles by their fields, execution %s', (execution: Execution) => {
  const opts = { clock: CLOCK, execution }

  describe.each(KINDS)('%s', (kind) => {
    it("a style's fields over the style each cell shows; null removes a field; equal styles are one; empty cells of the rectangle get a cell", async () => {
      const s = await (await opener(kind))(book(), opts)
      const r = await s.patchCells([{ sheet: 'S', range: 'A1:C3', set: { style: { font: { bold: true } } } }])
      const styles = [...book().styles!, { font: { italic: true, color: '#C00000', bold: true } }, { font: { bold: true }, fill: { color: '#C6EFCE' } }]
      expect(r.styles).toEqual(styles)
      const cells = cellsOf((await s.getBook()).sheets[0])
      // A1 italic and bold; empty cells bold, a cell of the filled column bold and filled; equal styles share an entry
      expect([cells.A1.style, cells.B1.style, cells.B2.style, cells.C1.style, cells.C3.style]).toEqual([3, 0, 0, 4, 4])
      expect(Object.keys(cells).filter((a) => /^[A-C][1-3]$/.test(a))).toHaveLength(9)
      // a field removed: the style left with nothing is the plain style, which the column's fill does not show through
      await s.patchCells([{ sheet: 'S', cell: 'C1', set: { style: { font: { bold: null }, fill: null } } }])
      const next = await s.getBook()
      expect(next.styles![cellsOf(next.sheets[0]).C1.style!]).toEqual({})
      // a style by its number replaces the cell's whole
      await s.patchCells([{ sheet: 'S', range: 'B2:B3', set: { style: 2 } }])
      expect(cellsOf((await s.getBook()).sheets[0]).B3).toEqual({ style: 2 })
      expect(await s.getHash()).toBe(compute(await s.getBook(), opts).hash)
      await s.close()
    })

    it('whole columns and rows: their own style and format, the cells of theirs with one, the cells where a formatted row crosses them merged', async () => {
      const s = await (await opener(kind))(book(), opts)
      const r = await s.patchCells([{ sheet: 'S', range: 'C:D', set: { style: { font: { italic: true } }, numberFormat: '#,##0' } }])
      expect(r.layout).toEqual(['S'])
      const b = await s.getBook()
      const st = (i: number | undefined) => (i === undefined ? undefined : b.styles![i])
      const sh = expandSheet(b.sheets[0])
      expect([st(sh.columns!.C.style), sh.columns!.C.numberFormat, st(sh.columns!.D.style), sh.columns!.D.numberFormat]).toEqual([{ fill: { color: '#C6EFCE' }, font: { italic: true } }, '#,##0', { font: { italic: true } }, '#,##0'])
      // C3 has a format of its own: the column's new one; where the bold row 5 crosses, what the cell showed (the row's bold) and italic
      expect(sh.cells.C3.numberFormat).toBe('#,##0')
      expect([st(sh.cells.C5.style), st(sh.cells.D5.style)]).toEqual([{ font: { bold: true, italic: true } }, { font: { bold: true, italic: true } }])
      // a row across the formatted columns: its style, and where it crosses them the columns' looks under it
      await s.patchCells([{ sheet: 'S', range: '7:7', set: { style: { fill: { color: '#FFEE99' } } } }])
      const c = await s.getBook()
      expect(c.styles![cellsOf(c.sheets[0]).C7.style!]).toEqual({ fill: { color: '#FFEE99' }, font: { italic: true } })
      expect(cellsOf(c.sheets[0]).A7).toBeUndefined()
      expect(await s.getHash()).toBe(compute(await s.getBook(), opts).hash)
      await s.close()
    })

    it('the number format by numberFormat, its old name refused; whole lines take no meta; a rectangle past the bound is refused', async () => {
      const s = await (await opener(kind))(book(), opts)
      const r = await s.patchCells([{ sheet: 'S', cell: 'B2', set: { numberFormat: '0.0' } }, { sheet: 'S', range: 'D5:D6', set: { numberFormat: '0.000', meta: { id: 1 } } }])
      expect(r.changes.map((c) => [c.cell, c.text, c.numberFormat])).toEqual([['B2', '2.0', '0.0'], ['D5', '4.000', '0.000'], ['D6', '', '0.000']])
      await expect(Promise.resolve(s.patchCells([{ sheet: 'S', range: 'B:B', set: { meta: 1 } }]))).rejects.toThrow('patches[0].set.meta: whole columns take no meta')
      await expect(Promise.resolve(s.patchCells([{ sheet: 'S', cell: 'B2', set: { format: '0', numberFormat: '0.0' } as never }]))).rejects.toThrow('patches[0].set: unknown field "format" (expected style, numberFormat, meta, note, link) (format is the old name of numberFormat)')
      await expect(Promise.resolve(s.patchCells([{ sheet: 'S', range: 'A1:Z50000', set: { style: { font: { bold: true } } } }]))).rejects.toThrow('formatting applies to at most 1048576 cells of a rectangle')
      await expect(Promise.resolve(s.patchCells([{ sheet: 'S', range: 'A1:B2', set: { style: { font: { size: 500 } } } }]))).rejects.toThrow('patches[0].set.style.font.size')
      // whole lines lose meta, notes and links of their cells with null
      await s.patchCells([{ sheet: 'S', range: '5:6', set: { meta: null } }])
      expect((await s.getBook()).sheets[0].cells.D5.meta).toBeUndefined()
      expect(s.revision).toBe(2)
      await s.close()
    })
  })
})
