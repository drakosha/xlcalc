/**
 * `readCells` gives each range's cells by rows, whatever order the book's cells were given in: input cells, cells of a
 * look alone and cells an array spills into together, a cell once in the first range that has it, the rows asked,
 * and pages by the cursor in that order. In a synchronous session, in the same thread and in a thread, with closures
 * and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { BookInput, Execution, ReadCellsResult } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-16T12:00:00Z', seed: 1 }

// the keys out of order; B5 and C5 spilled from A5; D1 a look alone; B4 an input cell after the spill's row in the keys
const book: BookInput = {
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: { C3: { value: 3 }, A5: { formula: 'SEQUENCE(1,3)' }, B1: { value: 'b' }, A3: { value: 'a' }, B4: { value: 4 }, A1: { value: 1 } },
      looks: [{ range: 'D1', style: 0 }],
    },
  ],
}

const cells = (p: ReadCellsResult) => p.cells.map((c) => c.cell)

describe.each(['closures', 'eval'] as Execution[])('readCells by rows, execution %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('each range by rows: input cells, looks and spills together, a cell once; the rows asked; pages in that order', async () => {
      const s = await (await opener(kind))(book, { clock, execution })
      expect(cells(await s.readCells([{ sheet: 'S', range: 'A1:D9' }], { fields: [] }))).toEqual(['A1', 'B1', 'D1', 'A3', 'C3', 'B4', 'A5', 'B5', 'C5'])
      // two ranges: the second's cells by rows after the first's, those the first has left out
      expect(cells(await s.readCells([{ sheet: 'S', range: 'B1:D5' }, { sheet: 'S', range: 'A1:D5' }], { fields: [] }))).toEqual(['B1', 'D1', 'C3', 'B4', 'B5', 'C5', 'A1', 'A3', 'A5'])
      expect(cells(await s.readCells([{ sheet: 'S', range: 'A1:XFD1048576' }], { rows: [3, 5], fields: [] }))).toEqual(['A3', 'C3', 'A5', 'B5', 'C5'])
      const all: string[] = []
      let cursor: string | undefined
      do {
        const p = await s.readCells([{ sheet: 'S', range: 'A1:D9' }], { maxCells: 2, count: true, fields: [], ...(cursor ? { cursor } : {}) })
        expect(p.total).toBe(9)
        all.push(...cells(p))
        cursor = p.cursor
      } while (cursor)
      expect(all).toEqual(['A1', 'B1', 'D1', 'A3', 'C3', 'B4', 'A5', 'B5', 'C5'])
      await s.close()
    })
  })
})
