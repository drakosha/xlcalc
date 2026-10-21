/**
 * Reproductions from the review of 21.10 on the pool of threads: a book whose text cannot be written, a restore that
 * cannot be sent, a thread that cannot start, and the deadline and signal of reads in a thread.
 */
import { describe, expect, it } from 'vitest'
import { createInlineCalc, type BookInput, type Calc, type CellInput } from '../src/index.js'
import type { Spawn } from '../src/worker/platform.js'
import { chainBook, fails, setA1 } from './async-books.js'
import { closedAfter, inside, pool } from './fixtures/pool.js'

const { keep, made } = closedAfter()

/** What `p` settles to, or 'hung' when it has not within `ms`. */
const settled = (p: Promise<unknown>, ms = 3000): Promise<unknown> =>
  Promise.race([p.then((v) => ['resolved', v], (e) => ['rejected', e]), new Promise((resolve) => setTimeout(() => resolve('hung'), ms))])

/** A book of `n` cells whose first cell's meta throws when it is read, as a cyclic one overflows the stack. */
function throwingBook(n: number): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= n; r++) cells[`A${r}`] = { value: r }
  cells.A1 = {
    value: 1,
    meta: {
      get x(): never {
        throw new Error('getter boom')
      },
    },
  } as unknown as CellInput
  return { sheets: [{ name: 'S', cells }] }
}

/** Over the 16,384 cells written at once: the book's text is written between turns of the event loop. */
const LARGE = 16_385

describe.each(['closures', 'eval'] as const)('the pool, %s', (execution) => {
  it('a book whose text cannot be written fails its call, and the tasks behind it run', async () => {
    const p = await made({ threads: 1 })
    expect(await settled(p.compute(throwingBook(10), { execution }))).toEqual(['rejected', expect.objectContaining({ message: 'getter boom' })])
    const bad = p.compute(throwingBook(LARGE), { execution })
    const next = p.compute(chainBook(2), { execution })
    expect(await settled(bad)).toEqual(['rejected', expect.objectContaining({ message: 'getter boom' })])
    expect(await settled(next)).toEqual(['resolved', expect.anything()])
    const opening = p.createSession(throwingBook(LARGE), { execution })
    const after = p.createSession(chainBook(2), { execution })
    const job = p.compute(chainBook(2), { execution })
    expect(await settled(opening)).toEqual(['rejected', expect.objectContaining({ message: 'getter boom' })])
    expect(await settled(after)).toEqual(['resolved', expect.anything()])
    expect(await settled(job)).toEqual(['resolved', expect.anything()])
    // refused before it waits: its text, still written, fails unseen
    const closed = await pool({ threads: 1 })
    await closed.close()
    await expect(closed.compute(throwingBook(LARGE), { execution })).rejects.toMatchObject({ code: 'closed' })
    await new Promise((resolve) => setTimeout(resolve, 100))
  }, 30_000)

  it('a restore that cannot be sent closes its session with restore_failed, once, and the thread goes on', async () => {
    const { calc, inside: x } = await inside({ threads: 1, maxSessionsPerThread: 1 })
    keep(calc)
    const a = await calc.createSession(chainBook(2), { execution })
    // b takes the full thread: a leaves it, to come back from its snapshot on its next call
    const b = await calc.createSession(chainBook(3), { execution })
    const ra = [...x.recs][0]
    ra.snapshot = { ...ra.snapshot, book: { sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { f() {} } } } }] } as never }
    const t = x.slots[0].thread!
    const post = t.post
    let restores = 0
    // the old pool posted the same restore for ever: past 50 the post goes through, so that the test ends
    t.post = (m) => post(m.t === 'restore' && ++restores > 50 ? { ...m, snapshot: { ...m.snapshot, book: chainBook(2) as never } } : m)
    const e = await fails(a.getCell('S', 'B1'))
    expect(restores).toBe(1)
    expect(e.code).toBe('restore_failed')
    expect(e.message).toMatch(/cannot be sent to a thread/)
    expect((await fails(a.setCells(setA1(3)))).code).toBe('restore_failed')
    expect((await b.getCell('S', 'B3'))?.value).toBe(4)
  }, 30_000)

  it('a thread that cannot start fails one task at a time, out of a lost thread and a restore too', async () => {
    let failing = false
    const wrap = (s: Spawn): Spawn => (m, l) => {
      if (failing) throw Object.assign(new Error('The Worker initialization failed: EAGAIN'), { code: 'ERR_WORKER_INIT_FAILED' })
      return s(m, l)
    }
    const { calc, crash } = await inside({ threads: 1 }, wrap)
    keep(calc)
    const busy = calc.compute(chainBook(2), { execution })
    const queued = [calc.compute(chainBook(2), { execution }), calc.compute(chainBook(2), { execution })]
    failing = true
    expect(() => crash(0)).not.toThrow()
    expect((await fails(busy)).code).toBe('thread_failed')
    for (const q of queued) {
      const r = await settled(q)
      expect(r).toEqual(['rejected', expect.objectContaining({ code: 'thread_failed', message: expect.stringMatching(/could not start.*EAGAIN/) })])
    }
    failing = false
    expect(await settled(calc.compute(chainBook(2), { execution }))).toEqual(['resolved', expect.anything()])

    // a session whose thread was lost: its call fails while no thread starts, and changes nothing
    const s = await calc.createSession(chainBook(3), { execution })
    crash(0)
    failing = true
    expect((await fails(s.setCells(setA1(2)))).code).toBe('thread_failed')
    failing = false
    expect((await s.getCell('S', 'A1'))?.value).toBe(1)
    expect(s.revision).toBe(0)
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect((await s.setCells(setA1(5))).changes.map((c) => c.cell)).toEqual(['A1', 'B1', 'B2', 'B3'])
  }, 30_000)
})

/** 5 × 20,000 values: reading them all takes far longer than a millisecond, in slices. */
const wide = (): BookInput => {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 20_000; r++) for (const c of 'ABCDE') cells[`${c}${r}`] = { value: r }
  return { sheets: [{ name: 'S', cells }] }
}

const CALCS = [
  ['in the same thread', (timeoutMs?: number): Promise<Calc> => Promise.resolve(keep(createInlineCalc({ sliceMs: 5, timeoutMs })))],
  ['in a thread', (timeoutMs?: number): Promise<Calc> => made({ threads: 1, sliceMs: 5, timeoutMs })],
] as const

describe.each(['closures', 'eval'] as const)('reads of a session %s', (execution) => {
  describe.each(CALCS)('%s', (_, make) => {
    it('getCells stops at its deadline and at its signal', async () => {
      const calc = await make()
      const s = await calc.createSession(wide(), { execution })
      expect((await fails(s.getCells('S', 'A1:E20000', { timeoutMs: 1 }))).code).toBe('soft_timeout')
      const ac = new AbortController()
      setTimeout(() => ac.abort(), 1)
      expect((await fails(s.getCells('S', 'A1:E20000', { signal: ac.signal }))).code).toBe('aborted')
      expect(Object.keys(await s.getCells('S', 'A1:B2', { timeoutMs: 5000 }))).toEqual(['A1', 'B1', 'A2', 'B2'])
    }, 30_000)

    it("the calculator's timeoutMs bounds the reads that take no options", async () => {
      const calc = await make(0)
      const s = await calc.createSession(chainBook(2), { execution, timeoutMs: 60_000 })
      for (const read of [() => s.getCells('S', 'A1:B2'), () => s.getCell('S', 'B1'), () => s.getViewport('S', 'A1:B2'), () => s.getUsedRange('S'), () => s.getEnd('S', 'A1', 'down'), () => s.getCellIssue('S', 'B1'), () => s.getValidation('S', 'B1')])
        expect((await fails(read())).code).toBe('soft_timeout')
      // a call's own deadline overrides it
      expect((await s.getCells('S', 'B1', { timeoutMs: 5000 })).B1).toMatchObject({ value: 2 })
    }, 30_000)
  })
})
