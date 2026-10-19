/**
 * `limits.writes`: the cells one edit of a session writes with a value or a formula, counted before anything is written.
 * Each edit it bounds (a copy, from and onto a filtered sheet and over part of an array formula, a move, fillCellText,
 * pasteText, replaceInCells, setCells, replaceCells, setCellTexts, applyChanges, and the previews) is refused one cell
 * over the limit, with the book, its revision, its history and its hash as before, and made at the limit as a session
 * without limits makes it; in a synchronous session, in the same thread and in a thread, with closures and eval. A copy
 * of one cell onto a whole column is refused within 50 ms; the calculators' limits hold under a session's own; what
 * writes no more than the sheet holds is not bounded.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, createInlineCalc, createSession, type AsyncSession, type BookInput, type CalcLimits, type Execution } from '../src/index.js'
import { pool } from './fixtures/pool.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** S: values, a formula, a look alone, an old-kind array and texts to replace; F: a filtered sheet, row 2 hidden under a spill. */
const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1 },
        A2: { formula: 'A1*2' },
        A3: { numberFormat: '0.00' },
        A4: { value: 'x', note: { text: 'n' } },
        M1: { formula: '{1;2;3}', array: 'M1:M3' },
        K1: { value: 'ax' },
        K2: { value: 'bx' },
        K3: { value: 'c' },
      },
    },
    { name: 'F', filtered: true, rows: { 2: { hidden: true } }, cells: { A1: { formula: 'SEQUENCE(3)' }, B1: { value: 'b1' }, B2: { value: 'b2' }, B3: { value: 'b3' }, C2: { value: 'c2' } } },
  ],
})

const ops = [
  { op: 'setCells' as const, edits: [{ sheet: 'S', cell: 'P1', input: { value: 1 } }, { sheet: 'S', cell: 'P2', input: { formula: 'P1+1' } }] },
  { op: 'replaceCells' as const, edits: [{ sheet: 'S', cell: 'P3', input: { value: 3 } }] },
  { op: 'patchCells' as const, patches: [{ sheet: 'S', range: 'Q1:Q100', set: { numberFormat: '0.0' } }] },
]

/** Each bounded edit, the cells it writes, and the call. */
const CASES: [string, number, (s: AsyncSession) => Promise<unknown>][] = [
  // A1, A2 and A4 hold content, A3 a look alone; two tiles
  ['a copy of a range onto two tiles', 6, (s) => s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1:A4' }, to: { sheet: 'S', range: 'C1:C8' } })],
  ['a copy of part of an array formula, each cell its own array', 2, (s) => s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'M2:M3' }, to: { sheet: 'S', range: 'N1' } })],
  // rows 1 and 3 as values: the spill's anchor and its third cell, B1 and B3; not B2 and C2 of the hidden row
  ['a copy from a filtered sheet', 4, (s) => s.editStructure({ op: 'copyRange', from: { sheet: 'F', range: 'A1:C3' }, to: { sheet: 'S', range: 'E1' } })],
  // the runs of visible rows 1 and 3:5 take one tile and three
  ['a copy onto a filtered sheet', 4, (s) => s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'F', range: 'D1:D5' } })],
  ['a move', 3, (s) => s.editStructure({ op: 'moveRange', from: { sheet: 'S', range: 'A1:A4' }, to: { sheet: 'S', cell: 'E1' } })],
  ['fillCellText', 10, (s) => s.fillCellText('S', 'G1:H5', 'x')],
  ['previewFillCellText', 10, (s) => s.previewFillCellText('S', 'G1:H5', '=ROW()')],
  // two rows, the second filled out to the first's width
  ['pasteText', 4, (s) => s.pasteText('S', 'J1', 'a\tb\nc')],
  ['previewPasteText', 4, (s) => s.previewPasteText('S', 'J1', 'a\tb\nc')],
  ['replaceInCells', 2, (s) => s.replaceInCells([{ sheet: 'S', range: 'K:K' }], 'x', 'y')],
  ['setCells', 3, (s) => s.setCells([1, 2, 3].map((r) => ({ sheet: 'S', cell: `P${r}`, input: { value: r } })))],
  ['replaceCells', 2, (s) => s.replaceCells([1, 2].map((r) => ({ sheet: 'S', cell: `P${r}`, input: { formula: `${r}*2` } })))],
  ['setCellTexts', 2, (s) => s.setCellTexts([1, 2].map((r) => ({ sheet: 'S', cell: `P${r}`, text: `${r}` })))],
  ['previewCellTexts', 2, (s) => s.previewCellTexts([1, 2].map((r) => ({ sheet: 'S', cell: `P${r}`, text: `${r}` })))],
  // the edits of cells of the batch; a patch is looks
  ['applyChanges', 3, (s) => s.applyChanges(ops)],
  ['previewChanges', 3, (s) => s.previewChanges(ops)],
]

/** What a refused call must leave as it was. */
const state = async (s: AsyncSession) => ({ book: JSON.stringify(await s.getBook()), revision: s.revision, history: s.history, hash: await s.getHash() })

const refusal = (e: unknown) => {
  expect((e as Error).name).toBe('CalcError')
  const { code, limit, max, size } = e as CalcError
  return [code, limit, max, size]
}

describe.each(['closures', 'eval'] as Execution[])('limits.writes, execution %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it.each(CASES)('%s: one cell over the limit is refused and changes nothing; at the limit it is made as without limits', async (_, n, call) => {
      const open = await opener(kind)
      const at = (writes?: number) => open(book(), { clock, execution, ...(writes ? { limits: { writes } } : {}) })
      const over = await at(n - 1)
      // a step in the history and a revision, which a refusal leaves
      await over.setCells([{ sheet: 'S', cell: 'Z1', input: { value: 0 } }])
      const before = await state(over)
      expect(refusal(await call(over).catch((e: unknown) => e))).toEqual(['limit', 'writes', n - 1, n])
      expect(await state(over)).toEqual(before)
      // the session goes on
      expect((await over.setCells([{ sheet: 'S', cell: 'Z1', input: { value: 1 } }])).revision).toBe(2)
      const [bounded, free] = [await at(n), await at()]
      expect(JSON.stringify(await call(bounded))).toBe(JSON.stringify(await call(free)))
      expect(await state(bounded)).toEqual(await state(free))
      await Promise.all([over, bounded, free].map((s) => s.close()))
    })

    it('a copy of one cell onto a whole column is refused within 50 ms, its value and its formula', async () => {
      const open = await opener(kind)
      for (const input of [{ value: 1 }, { formula: 'ROW()*2' }]) {
        const s = await open({ sheets: [{ name: 'S', cells: { A1: input } }] }, { clock, execution, limits: { writes: 100_000 } })
        const hash = await s.getHash()
        const t0 = performance.now()
        const e = await s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B:B' } }).catch((x: unknown) => x)
        const ms = performance.now() - t0
        expect(refusal(e)).toEqual(['limit', 'writes', 100_000, 1_048_576])
        expect(ms).toBeLessThan(50)
        expect([await s.getHash(), s.revision]).toEqual([hash, 0])
        await s.close()
      }
    })

    it("the refusal says it in the book's language, as a person's edit brings it", async () => {
      const open = await opener(kind)
      for (const [locale, message] of [[undefined, 'one edit writes 1048576 cells, over the limit of 100000'], ['ru-RU', 'Правка записывает слишком много ячеек: 1048576 при пределе 100000']] as const) {
        const s = await open({ ...(locale ? { locale } : {}), sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { clock, execution, limits: { writes: 100_000 } })
        const e = await s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B:B' } }).catch((x: unknown) => x)
        expect((e as CalcError).message).toBe(message)
        await s.close()
      }
    })
  })
})

describe('limits.writes', () => {
  it("a calculator's limit holds under a session's own limits of other fields, and a session's own takes its place", async () => {
    for (const calc of [createInlineCalc({ limits: { writes: 5 } }), await pool({ threads: 1, limits: { writes: 5 } })]) {
      const s = await calc.createSession(book(), { clock, limits: { formulas: 100 } })
      expect(refusal(await s.fillCellText('S', 'G1:H5', 'x').catch((e: unknown) => e))).toEqual(['limit', 'writes', 5, 10])
      const t = await calc.createSession(book(), { clock, limits: { writes: 10 } })
      expect(await t.fillCellText('S', 'G1:H5', 'x')).toMatchObject({ ok: true, revision: 1 })
      await calc.close()
    }
  })

  it('what writes no more than the sheet holds is not bounded: lines inserted and deleted, cells cleared, a copy of empty cells, looks', () => {
    const cells: BookInput['sheets'][0]['cells'] = {}
    for (let r = 1; r <= 50; r++) for (const c of 'ABCD') cells[`${c}${r}`] = { value: r }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, { clock, limits: { writes: 1 } })
    for (const op of [{ op: 'insertRows', sheet: 'S', row: 1, count: 3 }, { op: 'deleteRows', sheet: 'S', row: 1, count: 2 }, { op: 'insertColumns', sheet: 'S', column: 'A' }, { op: 'deleteColumns', sheet: 'S', column: 'B' }, { op: 'copyRange', from: { sheet: 'S', range: 'Z1:Z2' }, to: { sheet: 'S', range: 'A1:A100' } }] as const)
      expect(s.editStructure(op).ok, op.op).toBe(true)
    expect(s.patchCells([{ sheet: 'S', range: 'A1:D50', set: { numberFormat: '0.0' } }]).revision).toBe(6)
    expect(s.clearCells([{ sheet: 'S', range: 'A:D' }], 'all').revision).toBe(7)
    expect(s.getUsedRange('S')).toEqual({ rows: 0, columns: 0 })
  })

  it('is checked as input, a whole number of at least 1', () => {
    expect(() => createSession(book(), { limits: { writes: 0 } })).toThrow('opts.limits.writes: expected an integer from 1')
    expect(() => createSession(book(), { limits: { writes: 2.5 } as CalcLimits })).toThrow('opts.limits.writes: expected an integer from 1')
  })
})
