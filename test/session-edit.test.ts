/**
 * The edits of a session for a grid: `setCells` keeps a cell's look, `replaceCells` replaces the cell, `clearCells` as
 * Excel's ClearContents, ClearFormats and Clear; every edit answers its revision and what else changed (`layout`,
 * `styles`, `moved`, `repaint`), and `baseRevision` refuses an edit made over another revision. In a synchronous session,
 * in the same thread and in a thread, with closures and eval; undo takes each call back as one step.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, compute, InputError, type BookInput, type Execution } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { inside } from './fixtures/pool.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2, style: 0, format: '0.00', meta: { id: 7 }, note: { text: 'n' }, link: { url: 'https://example.com/' } },
        A2: { formula: 'A1*2', style: 1 },
        A3: { formula: 'SUM(A1:A2)' },
        C1: { value: 'Qty' },
        C2: { value: 5 },
      },
      columns: { E: { style: 1 } },
      tables: [{ name: 'T', range: 'C1:C2', columns: [{ name: 'Qty' }] }],
      conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['3'], style: { font: { bold: true } } }],
    },
    { name: 'R', cells: { A1: { formula: "SUM('S'!A:A)" } } },
  ],
})

describe.each(['closures', 'eval'] as const)('edits of a grid, execution %s', (execution: Execution) => {
  const opts = { clock: CLOCK, execution }

  describe.each(KINDS)('%s', (kind) => {
    it('setCells keeps what an edit leaves out; a value takes the formula\'s place; null clears the content and the hyperlink', async () => {
      const s = await (await opener(kind))(book(), opts)
      const r = await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
      expect(r.revision).toBe(1)
      expect(r.changes.map((c) => `${c.sheet}!${c.cell}=${JSON.stringify(c.value)} ${c.text}`)).toEqual(['S!A1=5 5.00', 'S!A2=10 10', 'S!A3=15 15', 'R!A1=30 30'])
      expect((await s.book()).sheets[0].cells.A1).toEqual({ value: 5, style: 0, format: '0.00', meta: { id: 7 }, note: { text: 'n' }, link: { url: 'https://example.com/' } })
      // a value over a formula: the formula goes, the style stays; the change carries the formula's end
      const f = await s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 1, style: null } }])
      expect(f.changes[0]).toEqual({ sheet: 'S', cell: 'A2', value: 1, text: '1', formula: null, style: null })
      expect((await s.book()).sheets[0].cells.A2).toEqual({ value: 1 })
      // a formula over a value: its canonical text in the change
      expect((await s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: '=a1*3' } }])).changes[0]).toMatchObject({ cell: 'A2', value: 15, formula: 'A1*3' })
      // Delete: the value and the hyperlink go, the style, format, meta and note stay
      await s.setCells([{ sheet: 'S', cell: 'A1', input: null }])
      expect((await s.book()).sheets[0].cells.A1).toEqual({ style: 0, format: '0.00', meta: { id: 7 }, note: { text: 'n' } })
      expect(s.revision).toBe(4)
      expect((await s.hash())).toBe(compute(await s.book(), opts).hash)
      await s.close()
    })

    it('replaceCells replaces a cell whole and removes it with null; the header of a table renames its column either way', async () => {
      const s = await (await opener(kind))(book(), opts)
      await s.replaceCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }, { sheet: 'S', cell: 'A2', input: null }])
      const cells = (await s.book()).sheets[0].cells
      expect([cells.A1, cells.A2]).toEqual([{ value: 5 }, undefined])
      const r = await s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 'Count' } }])
      expect(r.layout).toEqual(['S'])
      expect((await s.book()).sheets[0].tables![0].columns[0].name).toBe('Count')
      await s.close()
    })

    it('clearCells: contents keep the look, formats keep the content, all takes everything; whole lines lose their look; the rules of the cells shrink', async () => {
      const open = await opener(kind)
      const contents = await open(book(), opts)
      const c = await contents.clearCells([{ sheet: 'S', range: 'A1:A2' }], 'contents')
      expect(c.changes.map((x) => `${x.sheet}!${x.cell}=${JSON.stringify(x.value)}${'formula' in x ? ` ${x.formula}` : ''}`)).toEqual(['S!A1=null', 'S!A2=null null', 'S!A3=0', 'R!A1=0'])
      expect((await contents.book()).sheets[0].cells.A1).toEqual({ style: 0, format: '0.00', meta: { id: 7 }, note: { text: 'n' } })
      const formats = await open(book(), opts)
      const f = await formats.clearCells([{ sheet: 'S', range: 'A1:A3' }], 'formats')
      expect(f.changes.map((x) => [x.cell, x.text, x.style, x.numberFormat])).toEqual([['A1', '2', null, null], ['A2', '4', null, undefined]])
      // the rule over the cells cleared is gone, and the window paints them again
      expect(f.layout).toEqual(['S'])
      expect(f.repaint).toEqual([{ sheet: 'S', range: 'A1:A3' }])
      expect((await formats.layout()).sheets[0].conditionalFormats).toBeUndefined()
      const all = await open(book(), opts)
      await all.clearCells([{ sheet: 'S', range: 'A1' }, { sheet: 'S', range: 'E:E' }], 'all')
      const b = await all.book()
      expect([b.sheets[0].cells.A1, b.sheets[0].columns]).toEqual([undefined, undefined])
      expect(b.sheets[0].conditionalFormats).toEqual([{ type: 'cellIs', ranges: ['A2:A3'], operator: 'greaterThan', formulas: ['3'], style: { font: { bold: true } } }])
      for (const s of [contents, formats, all]) {
        expect(await s.hash()).toBe(compute(await s.book(), opts).hash)
        await s.close()
      }
    })

    it('an edit of another revision changes nothing and is refused with stale_revision; the revision of undo and redo', async () => {
      const s = await (await opener(kind))(book(), opts)
      expect(s.revision).toBe(0)
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { baseRevision: 0 })
      const before = JSON.stringify(await s.book())
      const e = await Promise.resolve(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 4 } }], { baseRevision: 0 })).catch((x) => x)
      expect(e).toBeInstanceOf(CalcError)
      expect((e as CalcError).code).toBe('stale_revision')
      expect(JSON.stringify(await s.book())).toBe(before)
      expect([s.revision, s.history]).toEqual([1, { undo: 1, redo: 0 }])
      // a window says the revision it is of
      expect((await s.getViewport('S', 'A1:A2')).revision).toBe(1)
      await expect(Promise.resolve(s.setCells([], { baseRevision: -1 }))).rejects.toThrow(InputError)
      expect((await s.undo({ baseRevision: 1 }))!.revision).toBe(2)
      expect((await s.redo())!.revision).toBe(3)
      // nothing to undo: no revision
      await s.undo()
      expect(await s.undo()).toBeNull()
      expect(s.revision).toBe(4)
      await s.close()
    })

    it('a structure operation answers how addresses moved, the layout and the formulas whose text changed; undo moves them back', async () => {
      const s = await (await opener(kind))(book(), opts)
      const r = await s.restructure({ op: 'insertRows', sheet: 'S', row: 2 })
      if (!r.ok) throw new Error(r.message)
      expect(r.moved).toEqual([{ kind: 'lines', sheet: 'S', axis: 'rows', at: 2, count: 1 }])
      expect(r.layout).toEqual(['S'])
      // A2's formula moves down unchanged; the sum over the rows reads one more
      expect(r.changes.filter((c) => c.formula !== undefined).map((c) => `${c.cell} ${c.formula}`)).toEqual(['A4 SUM(A1:A3)'])
      const renamed = await s.restructure({ op: 'renameSheet', sheet: 'R', name: 'Q' })
      expect(renamed.ok && renamed.moved).toEqual([{ kind: 'renameSheet', sheet: 'R', name: 'Q' }])
      expect((await s.undo())!.moved).toEqual([{ kind: 'renameSheet', sheet: 'Q', name: 'R' }])
      expect((await s.undo())!.moved).toEqual([{ kind: 'lines', sheet: 'S', axis: 'rows', at: 2, count: -1 }])
      await s.close()
    })

    it('setStyles answers the table; undo takes back each new call as one step', async () => {
      const s = await (await opener(kind))(book(), opts)
      const styles = [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }, { font: { italic: true } }]
      expect(await s.setStyles(styles)).toEqual({ revision: 1, changes: [], styles })
      const before = await s.book()
      await s.clearCells([{ sheet: 'S', range: 'A1:A3' }, { sheet: 'S', range: 'E:E' }], 'all')
      await s.patchCells([{ sheet: 'S', range: 'B1:B2', set: { style: { font: { italic: true } } } }])
      await s.replaceCells([{ sheet: 'S', cell: 'D9', input: { value: 1 } }])
      expect(s.history.undo).toBe(4)
      for (const call of ['replaceCells', 'patchCells', 'clearCells']) expect((await s.undo())!.call).toBe(call)
      expect(await s.book()).toEqual(before)
      await s.close()
    })
  })

  it('a thread lost: the session comes back at its revision, from the journal and from a snapshot, and refuses an edit over another', async () => {
    for (const snapshotEvery of [64, 2]) {
      const { calc, crash } = await inside({ threads: 1, snapshotEvery })
      const s = await calc.openSession(book(), opts)
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      await s.clearCells([{ sheet: 'S', range: 'A2' }], 'formats')
      await s.setCellText('S', 'B9', '1234567890')
      await s.patchCells([{ sheet: 'S', range: 'B1:B2', set: { style: { font: { italic: true } } } }])
      const [book4, revision] = [await s.book(), s.revision]
      expect(revision).toBe(4)
      crash(0)
      expect(await s.book()).toEqual(book4)
      await expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 9 } }], { baseRevision: 3 })).rejects.toMatchObject({ code: 'stale_revision' })
      expect((await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 9 } }], { baseRevision: 4 })).revision).toBe(5)
      // the history came back too: the five calls undone, one step each
      for (const call of ['setCells', 'patchCells', 'setCellText', 'clearCells', 'setCells']) expect((await s.undo())!.call).toBe(call)
      expect(await s.book()).toEqual(book())
      await calc.close()
    }
  }, 60_000)
})
