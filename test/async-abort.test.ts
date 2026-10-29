/**
 * An abort inside one long formula (`docs/async.md`): in a pool's thread the engine's checks read the pool's
 * cancellation in memory shared with the thread, so a heavy `SUMPRODUCT` stops within milliseconds of `abort()`, with
 * the outcome of an abort between formulas, and the thread goes on; a thread without that memory sees it when the
 * formula ends. On the same thread the checks read `signal.aborted`: a signal the platform fires waits for the formula,
 * one whose `aborted` reads memory another thread writes stops it.
 */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type BookInput } from '../src/index.js'
import { CLOCK, fails } from './async-books.js'
import { closedAfter, inside } from './fixtures/pool.js'
import type { ToThread } from '../src/worker/protocol.js'

/** One `SUMPRODUCT` of 16 million products in B1 once A1 is above 1: about half a second. */
const product = (a1 = 1): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: a1 }, B1: { formula: 'IF(A1>1,SUMPRODUCT(SEQUENCE(4000)*TRANSPOSE(SEQUENCE(4000))),A1)' }, B2: { formula: 'B1+1' } } }] })

/** How a call aborted `ms` after it starts settles, and in how many milliseconds. */
async function abortedAfter(ms: number, call: (signal: AbortSignal) => Promise<unknown>): Promise<{ code: string; cause: unknown; ms: number }> {
  const ac = new AbortController()
  setTimeout(() => ac.abort('stop'), ms)
  const t0 = performance.now()
  const e = await fails(call(ac.signal))
  return { code: e.code, cause: e.cause, ms: performance.now() - t0 }
}

/** The time of one computation of the heavy book by `run`, the machine as fast as it runs now. */
async function fullMs(run: (b: BookInput) => Promise<unknown>): Promise<number> {
  const t0 = performance.now()
  await run(product(2))
  return performance.now() - t0
}

const { made, keep } = closedAfter()

describe('an abort inside one long formula', () => {
  it("in a pool's thread: a heavy SUMPRODUCT stops within milliseconds, as an abort between formulas does, and the thread goes on", async () => {
    const p = await made({ threads: 1 })
    const full = await fullMs((b) => p.compute(b, { clock: CLOCK }))
    const stopped = await abortedAfter(50, (signal) => p.compute(product(2), { signal }))
    expect([stopped.code, stopped.cause]).toEqual(['aborted', 'stop'])
    expect(stopped.ms).toBeLessThan(Math.min(full / 2, 500))
    expect((await p.compute(product(), { clock: CLOCK })).hash).toBe(compute(product(), { clock: CLOCK }).hash)
  }, 30_000)

  it("a session's edit aborted inside a formula changes nothing, and the next edit is computed", async () => {
    const p = await made({ threads: 1 })
    const full = await fullMs((b) => p.compute(b, { clock: CLOCK }))
    const s = await p.createSession(product(), { clock: CLOCK })
    const before = (await s.getResult()).hash
    const stopped = await abortedAfter(50, (signal) => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }], { signal }))
    expect([stopped.code, stopped.cause]).toEqual(['aborted', 'stop'])
    expect(stopped.ms).toBeLessThan(Math.min(full / 2, 500))
    expect((await s.getBook()).sheets[0].cells.A1).toEqual({ value: 1 })
    expect((await s.getResult()).hash).toBe(before)
    const y = createSession(product(), { clock: CLOCK })
    const edit = [{ sheet: 'S', cell: 'A1', input: { value: 0.5 } }]
    expect(await s.setCells(edit)).toEqual(y.setCells(edit))
  }, 30_000)

  it('a thread without shared memory sees the abort when the formula ends, between its steps', async () => {
    // the thread gets no shared word, as in a browser page that is not cross-origin isolated
    const { calc } = await inside({ threads: 1 }, (spawn) => (onMessage, onLoss) => {
      const t = spawn(onMessage, onLoss)
      return { ...t, post: (m: ToThread) => t.post(m.t === 'init' ? { ...m, cancelled: undefined } : m) }
    })
    keep(calc)
    const full = await fullMs((b) => calc.compute(b, { clock: CLOCK }))
    const stopped = await abortedAfter(50, (signal) => calc.compute(product(2), { signal }))
    expect([stopped.code, stopped.cause]).toEqual(['aborted', 'stop'])
    expect(stopped.ms).toBeGreaterThan(full / 2)
  }, 30_000)

  it("on the same thread: the platform's signal waits for the formula, one that reads another thread's memory stops it", async () => {
    const calc = keep(createInlineCalc())
    const full = await fullMs((b) => calc.compute(b, { clock: CLOCK }))
    // an AbortController's abort is a task of the event loop, which runs when the formula ends
    const late = await abortedAfter(50, (signal) => calc.compute(product(2), { signal }))
    expect(late.code).toBe('aborted')
    expect(late.ms).toBeGreaterThan(full / 2)
    // a host on its own thread: the signal reads a word another thread sets 50 ms after the call starts
    const { Worker } = (await import('node:worker_threads')) as typeof import('node:worker_threads')
    const word = new Int32Array(new SharedArrayBuffer(4))
    const signal = { get aborted() { return Atomics.load(word, 0) === 1 }, reason: 'stop', addEventListener() {}, removeEventListener() {} }
    const t0 = performance.now()
    const setter = new Worker('const { workerData } = require("node:worker_threads"); setTimeout(() => Atomics.store(workerData, 0, 1), 50)', { eval: true, workerData: word })
    try {
      const e = await fails(calc.compute(product(2), { signal }))
      expect([e.code, e.cause]).toEqual(['aborted', 'stop'])
      expect(performance.now() - t0).toBeLessThan(Math.min(full / 2, 500))
    } finally {
      await setter.terminate()
    }
  }, 30_000)
})
