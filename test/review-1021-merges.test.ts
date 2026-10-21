/**
 * Merged areas, from the review of 21.10: `setLayout` merges over a table, a merge inside a merged area of entire
 * columns, and a session merging entire columns of a sheet that holds a single cell. In each case a session's
 * `getBook()` computes, and the session's hash is a full computation's.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, InputError, type AsyncSession, type BookInput, type EditResult, type Execution, type StructureOp } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

const tableBook = (): BookInput => ({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 'H1' }, B1: { value: 'H2' }, A2: { value: 6 }, A3: { value: 7 }, D1: { formula: 'SUM(T[H1])' } },
    tables: [{ name: 'T', range: 'A1:B3', columns: [{ name: 'H1' }, { name: 'H2' }] }],
  }],
})

describe.each(EXECUTIONS)('setLayout merges over a table (x-fuzz-session-8), %s', (execution) => {
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('refuses a merged area over a table with its path, and keeps one beside it', async () => {
      const s = await open(tableBook(), { execution, clock })
      const refused = s.setLayout([{ sheet: 'S', columns: 'D', set: { width: 12 } }, { sheet: 'S', merges: ['D5:E6', 'B3:C4'] }])
      await expect(refused).rejects.toThrow(InputError)
      await expect(refused).rejects.toThrow('edits[1].merges[1]: the merged area B3:C4 overlaps the table T')
      expect(s.revision).toBe(0)
      await s.setLayout([{ sheet: 'S', merges: ['D5:E6'] }])
      const after = await s.getBook()
      expect(after.sheets[0].merges).toEqual(['D5:E6'])
      expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
      await s.close()
    })
  })
})

const oneCell = (): BookInput => ({ styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }], sheets: [{ name: 'S', cells: { B6: { value: 'x y', style: 1 } } }] })

/** A sheet of one styled cell whose columns A:H are merged: its looks over the whole merged area. */
function mergedColumns(): BookInput {
  const r = editStructure(oneCell(), { op: 'mergeCells', sheet: 'S', range: 'A:H' })
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe.each(EXECUTIONS)('a merge inside a merged area of entire columns (x-fuzz-session-3), %s', (execution) => {
  const sheetOf = (b: BookInput) => ({ merges: b.sheets[0].merges, looks: b.sheets[0].looks, cells: b.sheets[0].cells })
  it.each([
    ['D:H', 'H'],
    ['A1:B2', 'H'],
    ['H16:K21', 'K'],
  ])('mergeCells %s grows to the area to column %s without putting its looks into cells', (range, right) => {
    const book = mergedColumns()
    expect(sheetOf(book)).toEqual({ merges: ['A1:H1048576'], looks: [{ range: 'B1:H1', style: 1 }, { range: 'A2:H1048576', style: 1 }], cells: { A1: { value: 'x y', style: 1 } } })
    const r = editStructure(book, { op: 'mergeCells', sheet: 'S', range })
    if (!r.ok) throw new Error(r.message)
    expect(sheetOf(r.book)).toEqual({ merges: [`A1:${right}1048576`], looks: [{ range: `B1:${right}1`, style: 1 }, { range: `A2:${right}1048576`, style: 1 }], cells: { A1: { value: 'x y', style: 1 } } })
    compute(r.book, { execution, clock })
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('merges inside the area, and undoes the merge', async () => {
      const s = await open(mergedColumns(), { execution, clock })
      const op: StructureOp = { op: 'mergeCells', sheet: 'S', range: 'D:H' }
      expect(await s.editStructure(op)).toMatchObject({ ok: true, changes: [] })
      const after = await s.getBook()
      expect(after.sheets[0].merges).toEqual(['A1:H1048576'])
      expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
      await s.undo()
      expect([s.revision, (await s.getBook()).sheets[0].merges]).toEqual([2, ['A1:H1048576']])
      await s.close()
    })
  })
})

describe.each(EXECUTIONS)('a session merging entire columns of a sheet of one cell (x-fuzz-session-4), %s', (execution) => {
  /** The session's hash a full computation's, and a window inside the rectangles to repaint as a session opened on its book shows it. */
  async function sameAsFull(s: AsyncSession, r: EditResult, repaint: string, display: boolean) {
    expect(r.repaint).toEqual([{ sheet: 'S', range: repaint }])
    const book = await s.getBook()
    expect(compute(book, { execution, clock }).hash).toBe(await s.getHash())
    const full = createSession(book, { execution, clock, display }).getViewport('S', 'C99999:D100000')
    const view = await s.getViewport('S', 'C99999:D100000')
    expect([view.cells, view.styles]).toEqual([full.cells, full.styles])
  }

  /** The result of a call that went through. */
  const done = (r: object | null): EditResult => {
    if (!r || ('ok' in r && r.ok === false)) throw new Error(`refused: ${JSON.stringify(r)}`)
    return r as EditResult
  }

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it.each([false, true])('display %s: the merge, its undo and redo list the first rows by cell and repaint the rest', async (display) => {
      const s = await open(oneCell(), { execution, clock, display })
      const t = performance.now()
      const r = done(await s.editStructure({ op: 'mergeCells', sheet: 'S', range: 'A:H' }))
      // listed by cell, the call took 10-16 s and returned 8,388,607 changes
      expect(performance.now() - t).toBeLessThan(1500)
      expect(r.changes.map((c) => [c.cell, c.style])).toEqual(['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1'].map((a) => [a, 1]))
      await sameAsFull(s, r, 'A2:H1048576', display)
      expect((await s.getViewport('S', 'C99999')).cells.C99999.style).toBe(1)
      const undo = done(await s.undo())
      // the rows above the moved cell are listed, those below it repainted
      expect(undo.changes.length).toBe(6 * 8 - 2)
      await sameAsFull(s, undo, 'A7:H1048576', display)
      expect((await s.getViewport('S', 'C99999')).cells.C99999).toBeUndefined()
      const redo = done(await s.redo())
      expect(redo.changes).toEqual(r.changes)
      await sameAsFull(s, redo, 'A2:H1048576', display)
      await s.close()
    })
  })
})
