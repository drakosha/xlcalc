/**
 * `patchCells` and `clearCells` over large rectangles of a sheet's looks (`SheetInput.looks`): the looks change by
 * pieces, not cell by cell, so formatting whole columns of them takes milliseconds; a rectangle of more than 65,536
 * cells whose look changed comes whole as a range of `repaint`, the cells of entries and the smaller rectangles as
 * `changes`; one step of the undo puts it back. In every kind of session, with closures and eval, display on and off.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, type AsyncSession, type BookInput, type EditResult, type Execution, type SessionOperation, type Viewport } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const styles = [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }]

/** One value, a spill and a styled value over a million rows of a filled look: before 22.10 a call over them took 40 s. */
const book = (): BookInput => ({
  styles,
  sheets: [{ name: 'S', cells: { A1: { value: 1 }, E1: { formula: 'SEQUENCE(3)' }, B5: { value: 2, style: 1 } }, looks: [{ range: 'A2:H1048576', style: 1 }] }],
})

/** The calls over the looks, each with the range it repaints and the cells it lists besides those of rows 2 to 5 (display on). */
const CALLS: [string, (s: AsyncSession) => Promise<EditResult>, string[], string[]][] = [
  ['patchCells on whole columns', (s) => s.patchCells([{ sheet: 'S', range: 'A:H', set: { style: 0, numberFormat: '0.00' } }]), ['A6:H1048576'], ['A1', 'E1']],
  ['patchCells on a rectangle', (s) => s.patchCells([{ sheet: 'S', range: 'A2:H131072', set: { style: { font: { italic: true } } } }]), ['A6:H131072'], []],
  ['clearCells of formats', (s) => s.clearCells([{ sheet: 'S', range: 'A2:H1048576' }], 'formats'), ['A6:H1048576'], []],
  ['applyChanges of a patch', (s) => s.applyChanges([{ op: 'patchCells', patches: [{ sheet: 'S', range: 'A:H', set: { style: 0 } }] } satisfies SessionOperation]), ['A6:H1048576'], []],
]

/** The rows 2 to 5 of the looks, which B5 cuts into rectangles of a few cells, listed by cell, and B5 itself. */
const AROUND = ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5']

/** A book with each style given by its fields, not its number: the table's new styles may come in another order. */
function byContent(b: BookInput): unknown {
  const look = <T extends { style?: number }>(x: T) => (x.style === undefined ? x : { ...x, style: b.styles![x.style] })
  const lines = (m: Record<string, { style?: number }> | undefined) => m && Object.fromEntries(Object.entries(m).map(([k, v]) => [k, look(v)]))
  return b.sheets.map((sh) => ({ ...sh, cells: lines(sh.cells), rows: lines(sh.rows), columns: lines(sh.columns), looks: sh.looks?.map(look) }))
}

const view = async (s: { getViewport(sheet: string, range: string): Viewport | Promise<Viewport> }, range: string) => {
  const { revision: _, ...v } = await s.getViewport('S', range)
  return v
}

describe.each(['closures', 'eval'] as const)('large looks, %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    for (const display of [false, true])
      it.each(CALLS)(`%s, display ${display}`, async (_, call, repaint, lines) => {
        const open = await opener(kind)
        const s = await open(book(), { clock: CLOCK, execution, display })
        // a synchronous session gives its live book, whose looks a call sets in place: compared as text taken before
        const before = JSON.stringify(await s.getBook())
        const t = performance.now()
        const r = await call(s)
        // the old way listed 8 million cells in 40 s
        expect(performance.now() - t).toBeLessThan(4000)
        expect(r.repaint).toEqual(repaint.map((range) => ({ sheet: 'S', range })))
        const cells = r.changes.map((c) => c.cell)
        // the cells of a value whose column's new format shows through, read where display is on
        expect(cells.filter((a) => !AROUND.includes(a))).toEqual(display ? lines : [])
        expect(AROUND.filter((a) => !cells.includes(a))).toEqual([])
        const after = JSON.parse(JSON.stringify(await s.getBook())) as BookInput
        // the session's state is that of its book, in the repainted rows too
        expect(await s.getHash()).toBe(compute(after, { clock: CLOCK, execution }).hash)
        const fresh = createSession(after, { clock: CLOCK, execution, display })
        for (const range of ['A1:H8', 'A131070:H131074', 'A1048570:H1048576']) expect(await view(s, range)).toEqual(await view(fresh, range))
        const back = (await s.undo())!
        expect(back.repaint).toEqual(r.repaint)
        expect(JSON.stringify(await s.getBook())).toBe(before)
        expect(await s.getHash()).toBe(compute(JSON.parse(before) as BookInput, { clock: CLOCK, execution }).hash)
        await s.redo()
        expect(JSON.parse(JSON.stringify(await s.getBook()))).toEqual(after)
        await s.close()
      })
  })
})

/** A book of formatted rows, columns and ranges, with values, a spill and a note among them. */
const formatted = (): BookInput => ({
  styles,
  sheets: [
    {
      name: 'S',
      columns: { F: { numberFormat: '0.000' } },
      cells: { A1: { value: 1 }, E1: { formula: 'SEQUENCE(3)' }, B5: { value: 2, style: 1 }, C7: { note: { text: 'n' } } },
      looks: [{ range: 'A2:H40000', style: 1 }],
    },
  ],
})

const patch = (range: string, set: object): SessionOperation => ({ op: 'patchCells', patches: [{ sheet: 'S', range, set }] }) as SessionOperation
const clear = (range: string, what: 'contents' | 'formats' | 'all'): SessionOperation => ({ op: 'clearCells', ranges: [{ sheet: 'S', range }], what })

/** Batches whose later operations read the looks the earlier ones set, over large ranges and whole lines. */
const BATCHES: [string, SessionOperation[]][] = [
  ['a style removed from columns, then a cell made italic', [patch('B:D', { style: null }), patch('B6', { style: { font: { italic: true } } })]],
  ['a style removed from columns, then a format over many rows', [patch('B:D', { style: null }), patch('A20000:H30000', { numberFormat: '0.0' })]],
  ['formats cleared from a column, then a format over many rows', [clear('C:C', 'formats'), patch('A20000:H30000', { numberFormat: '0.0' })]],
  ['rows made italic, then a value typed', [patch('2:40000', { style: { font: { italic: true } } }), { op: 'setCells', edits: [{ sheet: 'S', cell: 'A3', input: { value: 5 } }] }]],
  ['columns made italic, a cell replaced, a rectangle filled', [patch('A:H', { style: { font: { italic: true } } }), { op: 'replaceCells', edits: [{ sheet: 'S', cell: 'C3', input: null }] }, patch('B2:E30000', { style: { fill: { color: '#00FF00' } } })]],
  ['a cell emptied, then formatted twice by large rectangles', [{ op: 'replaceCells', edits: [{ sheet: 'S', cell: 'A11', input: null }] }, patch('A2:C30000', { style: { font: { italic: true } } }), patch('A11:F28000', { numberFormat: '0.00' })]],
  ['all cleared from a rectangle, then columns formatted', [clear('A3:H39000', 'all'), patch('C:F', { style: 0, numberFormat: '@' })]],
]

describe.each(['closures', 'eval'] as const)('applyChanges over large looks, %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it.each(BATCHES)('%s: the book of the calls one by one', async (_, ops) => {
      const open = await opener(kind)
      const [batch, single] = [await open(formatted(), { clock: CLOCK, execution }), await open(formatted(), { clock: CLOCK, execution })]
      const before = JSON.stringify(await batch.getBook())
      await batch.applyChanges(ops)
      for (const op of ops) await single.applyChanges([op])
      // as text taken at once (a synchronous session's book is live), compared field by field: cells may come in another order
      const book = JSON.parse(JSON.stringify(await batch.getBook())) as BookInput
      expect(book).toEqual(JSON.parse(JSON.stringify(await single.getBook())))
      expect(await batch.getHash()).toBe(compute(book, { clock: CLOCK, execution }).hash)
      await batch.undo()
      expect(JSON.parse(JSON.stringify(await batch.getBook()))).toEqual(JSON.parse(before))
      await batch.redo()
      expect(JSON.parse(JSON.stringify(await batch.getBook()))).toEqual(book)
      await Promise.all([batch.close(), single.close()])
    })
  })
})

describe('large looks, the cells by pieces', () => {
  it('a rectangle of more than 65,536 cells is formatted as its parts of fewer cells are, one by one', () => {
    const b: BookInput = {
      styles,
      sheets: [
        {
          name: 'S',
          rows: { '3': { style: 1 } },
          columns: { C: { numberFormat: '0.0' } },
          cells: { A1: { value: 1, style: 0 }, B7: { formula: 'SEQUENCE(4,2)' }, C40000: { value: 3, note: { text: 'n' } } },
          looks: [{ range: 'A2:D9', style: 1 }, { range: 'B30000:D50000', numberFormat: '0.00' }],
        },
      ],
    }
    for (const set of [{ style: { font: { italic: true } } }, { style: 0 }, { numberFormat: '@' }, { style: null, numberFormat: null }]) {
      const [whole, parts] = [createSession(b), createSession(b)]
      whole.patchCells([{ sheet: 'S', range: 'A1:D40000', set }])
      parts.patchCells([{ sheet: 'S', range: 'A1:D16000', set }, { sheet: 'S', range: 'A16001:D32000', set }, { sheet: 'S', range: 'A32001:D40000', set }])
      expect(byContent(whole.getBook())).toEqual(byContent(parts.getBook()))
      whole.clearCells([{ sheet: 'S', range: 'A2:D60000' }], 'formats')
      parts.clearCells([{ sheet: 'S', range: 'A2:D30000' }, { sheet: 'S', range: 'A30001:D60000' }], 'formats')
      expect(byContent(whole.getBook())).toEqual(byContent(parts.getBook()))
    }
  })

  it('formatting goes to the looks, which `limits.writes` does not count', () => {
    const open = () => createSession(book(), { limits: { writes: 10 } })
    expect(open().patchCells([{ sheet: 'S', range: 'A:H', set: { style: 0 } }]).repaint).toEqual([{ sheet: 'S', range: 'A6:H1048576' }])
    expect(open().clearCells([{ sheet: 'S', range: 'A2:H1048576' }], 'formats').repaint).toEqual([{ sheet: 'S', range: 'A6:H1048576' }])
  })
})
