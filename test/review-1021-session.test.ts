/**
 * Reproductions from the 21.10 review, session: a call or an undo stopped part way (its signal, its deadline, a limit)
 * changes nothing, and a clear of whole lines over a table's header keeps the book and the state together. After each
 * stop the session's hash is the full recalculation's of its book, the book passes the checks of a new one, and undo and
 * redo go on. In a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, compute, createInlineCalc, createSession, type AsyncSession, type BookInput, type Calc, type CalcOptions, type Execution, type Session } from '../src/index.js'
import { checkBook } from '../src/input.js'
import { CLOCK } from './async-books.js'
import { pool } from './fixtures/pool.js'

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))
const inline = () => {
  const c = createInlineCalc({ sliceMs: 0 })
  calcs.push(c)
  return c
}
let shared: Calc | undefined
const threads = async () => shared ?? (calcs.push((shared = await pool({ threads: 1 }))), shared)

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
})
