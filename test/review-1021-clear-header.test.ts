/**
 * Cells cleared through a table's header, from the review of 21.10 and the acceptance of r1/session: `applyChanges` with
 * `clearCells` renames the column as `clearCells` does, Clear All over a header on a styled row comes back whole after an
 * undo and a redo, and a range through a header clears the cells that hold a look alone. In every kind of session the
 * book stays one that computes to the session's hash.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type AsyncSession, type BookInput, type ClearKind, type Execution } from '../src/index.js'
import { checkBook } from '../src/input.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

/** Table Sales over A1:B2 with a formula on its Qty column, a bold first row, and cells below with a look and nothing else. */
const book = (): BookInput => ({
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'Data',
    cells: { A1: { value: 'Item' }, B1: { value: 'Qty' }, A2: { value: 'x' }, B2: { value: 2 }, D1: { formula: 'SUM(Sales[Qty])' }, B5: { style: 0 }, B6: { value: 6, style: 0 } },
    rows: { '1': { style: 0 } },
    tables: [{ name: 'Sales', range: 'A1:B2', columns: [{ name: 'Item' }, { name: 'Qty' }] }],
  }],
})

/** The session's book, which must be valid and compute to its hash. */
async function held(s: AsyncSession, execution: Execution): Promise<BookInput> {
  const b = structuredClone(await s.getBook())
  checkBook(structuredClone(b))
  expect(compute(b, { execution, clock }).hash).toBe(await s.getHash())
  return b
}

describe.each(EXECUTIONS)('%s', (execution) => {
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    const cleared = async (what: ClearKind, range: string, through: 'clearCells' | 'applyChanges') => {
      const s = await open(book(), { execution, clock })
      const ranges = [{ sheet: 'Data', range }]
      if (through === 'clearCells') await s.clearCells(ranges, what)
      else await s.applyChanges([{ op: 'clearCells', ranges, what }])
      return s
    }

    it('applyChanges clears a header cell as clearCells does (apply-clear-header)', async () => {
      for (const what of ['contents', 'all'] as const) {
        const [a, b] = [await cleared(what, 'B1', 'clearCells'), await cleared(what, 'B1', 'applyChanges')]
        const [x, y] = [await held(a, execution), await held(b, execution)]
        expect(y).toEqual(x)
        expect(x.sheets[0].tables![0].columns.map((c) => c.name)).toEqual(['Item', 'Column1'])
        expect(x.sheets[0].cells.D1.formula).toBe('SUM(Sales[Column1])')
        await a.close()
        await b.close()
      }
    })

    it('Clear All over a header on a styled row comes back whole after an undo and a redo (A)', async () => {
      const fresh = await open(book(), { execution, clock })
      const before = await held(fresh, execution)
      await fresh.close()
      for (const through of ['clearCells', 'applyChanges'] as const) {
        const s = await cleared('all', 'B1', through)
        const after = await held(s, execution)
        await s.undo()
        expect(await held(s, execution)).toEqual(before)
        await s.redo()
        expect(await held(s, execution)).toEqual(after)
        await s.close()
      }
    })

    it('a range through a header clears the cells with a look alone (B)', async () => {
      for (const through of ['clearCells', 'applyChanges'] as const) {
        for (const range of ['B1:B6', 'B:B']) {
          const s = await cleared('all', range, through)
          const b = await held(s, execution)
          expect([b.sheets[0].cells.B5, b.sheets[0].cells.B6, b.sheets[0].looks ?? []], `${through} ${range}`).toEqual([undefined, undefined, []])
          expect((await s.getCells('Data', 'B5')).B5, `${through} ${range}`).toBeUndefined()
          await s.close()
        }
      }
    })
  })
})
