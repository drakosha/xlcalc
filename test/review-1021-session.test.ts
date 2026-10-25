/**
 * Reproductions from the 21.10 review, session: a call or an undo stopped part way (its signal, its deadline, a limit)
 * changes nothing, and a clear of whole lines over a table's header keeps the book and the state together. After each
 * stop the session's hash is the full recalculation's of its book, the book passes the checks of a new one, and undo and
 * redo go on. In a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, compute, createInlineCalc, createSession, type AsyncSession, type BookInput, type Calc, type CalcOptions, type Execution, type Session } from '../src/index.js'
import { checkBook } from '../src/input.js'
import { CLOCK } from './async-books.js'
import { closedAfter } from './fixtures/pool.js'

const { keep, made } = closedAfter()
const inline = () => keep(createInlineCalc({ sliceMs: 0 }))
let shared: Calc | undefined
const threads = async () => shared ?? (shared = await made({ threads: 1 }))

/** A signal that reports aborted from its n-th poll on: a stop at each point of a call in turn. */
const stopsAt = (n: number) => {
  let polls = 0
  return { get aborted() { return ++polls >= n }, reason: 'stopped', addEventListener() {}, removeEventListener() {} } as unknown as AbortSignal
}

/** What a session holds: its book, hash and history. */
const held = async (s: Session | AsyncSession) => ({ book: structuredClone(await s.getBook()), hash: await s.getHash(), history: s.history })

/** The session's hash is the full recalculation's of its book, which a new session would take. */
async function consistent(s: Session | AsyncSession, execution: Execution) {
  const book = await s.getBook()
  expect(() => checkBook(book)).not.toThrow()
  expect(await s.getHash()).toBe(compute(book, { clock: CLOCK, execution }).hash)
}

/** The code of the error a call threw, none when it completed. */
const stopped = (p: Promise<unknown> | (() => unknown)) =>
  (typeof p === 'function' ? Promise.resolve().then(p) : p).then(
    () => null,
    (e) => {
      expect(e).toBeInstanceOf(CalcError)
      return (e as CalcError).code
    },
  )

// B2:C3 merged clears C2 and C3; A3's spill A3:B3 meets the area, so the merge (and its undo) computes it again after the cells
const mergeBook = (): BookInput => ({ sheets: [{ name: 'S', cells: { B2: { value: 1 }, C2: { value: 2 }, C3: { formula: 'B2+10' }, A3: { formula: 'SEQUENCE(1,2)' } } }] })
// the same with values to clear, which compute nothing, and a long chain on B3, which the spill's change computes again: a deadline lands there
const chainedMerge = (n = 60_000): BookInput => {
  const cells: BookInput['sheets'][number]['cells'] = { ...mergeBook().sheets[0].cells, C3: { value: 'y' }, E1: { formula: 'B3' } }
  for (let r = 2; r <= n; r++) cells[`E${r}`] = { formula: `E${r - 1}+1` }
  return { sheets: [{ name: 'S', cells }] }
}
const MERGE = { op: 'mergeCells', sheet: 'S', range: 'B2:C3' } as const

describe.each(['closures', 'eval'] as const)('a session stopped part way, %s', (execution: Execution) => {
  const o = { clock: CLOCK, execution }

  describe('fix-u311: a merge stopped while it computes the spills it meets', () => {
    it('in the same thread, stopped at each point: the cells it cleared come back, no step is kept', async () => {
      const calc = inline()
      let stops = 0
      for (let n = 1; ; n++) {
        const s = await calc.createSession(mergeBook(), o)
        const before = await held(s)
        const code = await stopped(s.editStructure(MERGE, { signal: stopsAt(n) }))
        if (!code) break
        stops++
        expect(code, `poll ${n}`).toBe('aborted')
        expect(await held(s), `poll ${n}`).toEqual(before)
        await consistent(s, execution)
        expect(await s.undo()).toBeNull()
        // the merge done after the stop is the one done at once
        await s.editStructure(MERGE)
        expect((await s.getBook()).sheets[0].merges).toEqual(['B2:C3'])
        await consistent(s, execution)
        await s.undo()
        expect(await held(s), `poll ${n}`).toEqual({ ...before, history: { undo: 0, redo: 1 } })
        await s.close()
      }
      expect(stops).toBeGreaterThan(3)
    })

    it('in a thread, by its deadline', async () => {
      const calc = await threads()
      for (const timeoutMs of [1, 10]) {
        const s = await calc.createSession(chainedMerge(), o)
        const before = await held(s)
        expect(await stopped(s.editStructure(MERGE, { timeoutMs }))).toBe('soft_timeout')
        expect(await held(s)).toEqual(before)
        await consistent(s, execution)
        await s.editStructure(MERGE)
        await s.undo()
        expect(await held(s)).toEqual({ ...before, history: { undo: 0, redo: 1 } })
        await s.close()
      }
    }, 60_000)
  })

  describe("fix-u312: an undo or a redo of a merge stopped in the merge's computation", () => {
    it('in the same thread, stopped at each point: the book, hash and history stay; undo and redo then give the books before', async () => {
      const calc = inline()
      let stops = 0
      for (const way of ['undo', 'redo'] as const) {
        for (let n = 1; ; n++) {
          const s = await calc.createSession(mergeBook(), o)
          const original = await held(s)
          await s.editStructure(MERGE)
          const merged = await held(s)
          if (way === 'redo') await s.undo()
          const before = await held(s)
          const code = await stopped(s[way]({ signal: stopsAt(n) }))
          if (!code) break
          stops++
          expect(await held(s), `${way} poll ${n}`).toEqual(before)
          await consistent(s, execution)
          if (way === 'undo') {
            await s.undo()
            expect((await held(s)).book).toEqual(original.book)
            await s.redo()
            expect(await held(s)).toEqual(merged)
          } else {
            await s.redo()
            expect(await held(s)).toEqual(merged)
            await s.undo()
            expect(await held(s)).toEqual(before)
          }
          await consistent(s, execution)
          await s.close()
        }
      }
      expect(stops).toBeGreaterThan(6)
    })

    it('in a thread, by its deadline', async () => {
      const calc = await threads()
      const s = await calc.createSession(chainedMerge(), o)
      const original = await held(s)
      await s.editStructure(MERGE)
      const merged = await held(s)
      expect(await stopped(s.undo({ timeoutMs: 1 }))).toBe('soft_timeout')
      expect(await held(s)).toEqual(merged)
      await consistent(s, execution)
      await s.undo()
      expect((await held(s)).book).toEqual(original.book)
      const undone = await held(s)
      expect(await stopped(s.redo({ timeoutMs: 1 }))).toBe('soft_timeout')
      expect(await held(s)).toEqual(undone)
      await consistent(s, execution)
      await s.redo()
      expect(await held(s)).toEqual(merged)
      await s.close()
    }, 60_000)
  })

  describe('fix-u305: a preview or a check of an entry stopped while it computes', () => {
    // A1 under a rule, and a long chain on it
    const chain = (n = 20_000): BookInput => {
      const cells: BookInput['sheets'][number]['cells'] = { A1: { value: 1 } }
      for (let r = 2; r <= n; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
      return { sheets: [{ name: 'S', cells, validations: [{ range: 'A1', type: 'whole', operator: 'between', formula1: '0', formula2: '10' }] }] }
    }
    type Stop = (s: AsyncSession, opts: CalcOptions) => Promise<unknown>
    const CALLS: [string, Stop][] = [
      ['validateEntry', (s, opts) => s.validateEntry('S', 'A1', { value: 5 }, opts)],
      ['previewCellText', (s, opts) => s.previewCellText('S', 'A1', '7', opts)],
    ]
    const limited = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] })

    /** After a preview over the spill limit the session goes on, with its book as before. */
    async function afterLimit(s: Session | AsyncSession) {
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }])
      const before = await held(s)
      expect(await stopped(() => s.previewCellText('S', 'C1', '=SEQUENCE(1000)'))).toBe('limit')
      expect(await held(s)).toEqual(before)
      await consistent(s, execution)
      expect((await s.undo())!.changes).toEqual([{ sheet: 'S', cell: 'A1', value: 1, text: '1' }])
      await s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 3 } }])
      expect((await s.getBook()).sheets[0].cells).toEqual({ A1: { value: 1 }, C1: { value: 3 } })
    }

    it('synchronous: a preview over a limit leaves the book and the session usable', async () => {
      await afterLimit(createSession(limited(), { ...o, limits: { spill: 100 } }))
    })

    it('in the same thread, by its deadline, its signal and a limit', async () => {
      const calc = inline()
      for (const [name, call] of CALLS) {
        for (const opts of [{ timeoutMs: 1 }, { signal: stopsAt(3) }] as CalcOptions[]) {
          const s = await calc.createSession(chain(), o)
          const before = await held(s)
          expect(await stopped(call(s, opts)), name).not.toBeNull()
          expect(await held(s), name).toEqual(before)
          expect((await s.getCells('S', ['A20000'])).A20000.value).toBe(20_000)
          await consistent(s, execution)
          await s.close()
        }
      }
      await afterLimit(await calc.createSession(limited(), { ...o, limits: { spill: 100 } }))
    }, 60_000)

    it('in a thread, by its deadline and a limit', async () => {
      const calc = await threads()
      for (const [name, call] of CALLS) {
        const s = await calc.createSession(chain(), o)
        const before = await held(s)
        expect(await stopped(call(s, { timeoutMs: 1 })), name).toBe('soft_timeout')
        expect(await held(s), name).toEqual(before)
        await consistent(s, execution)
        await s.close()
      }
      await afterLimit(await calc.createSession(limited(), { ...o, limits: { spill: 100 } }))
    }, 60_000)
  })

  describe('x-fuzz-session-7: clearing all of whole lines through a table header', () => {
    const book = (): BookInput => ({
      sheets: [
        { name: 'Data', cells: { A1: { value: 'Item' }, B1: { value: 'Qty' }, C1: { value: 'Price' }, A2: { value: 'x' }, B2: { value: 2 }, C2: { value: 3 }, A3: { value: 'y' }, B3: { value: 4 }, C3: { value: 5 } }, tables: [{ name: 'Sales', range: 'A1:C3', columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Price' }] }] },
        { name: 'Other', cells: { A1: { formula: 'COUNTA(Sales[Qty])' } } },
      ],
    })

    /** Clears of whole columns and of whole rows: the book holds what the session computes, and undo and redo go back and forth. */
    async function clears(s: () => Promise<Session | AsyncSession> | Session | AsyncSession) {
      for (const [range, header] of [['B:C', ['Item', 'Column1', 'Column2']], ['1:1', ['Column1', 'Column2', 'Column3']]] as const) {
        const x = await s()
        const original = await held(x)
        await x.clearCells([{ sheet: 'Data', range }], 'all')
        const cleared = await held(x)
        const data = (await x.getBook()).sheets[0]
        expect(data.tables![0].columns.map((c) => c.name), range).toEqual(header)
        expect(header.map((_, c) => data.cells[`${'ABC'[c]}1`]?.value)).toEqual(header)
        await consistent(x, execution)
        await x.undo()
        expect(await held(x), range).toEqual({ ...original, history: { undo: 0, redo: 1 } })
        await consistent(x, execution)
        await x.redo()
        expect(await held(x), range).toEqual(cleared)
        await x.close()
      }
    }

    it('synchronous', () => clears(() => createSession(book(), o)))
    it('in the same thread', () => clears(() => inline().createSession(book(), o)))
    it('in a thread', async () => {
      const calc = await threads()
      await clears(() => calc.createSession(book(), o))
    }, 60_000)
  })
})
