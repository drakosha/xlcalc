/**
 * The pool at its edges: a book or an edit that cannot cross to a thread, closing a session while the pool closes or
 * while its thread is lost, admission with `maxQueue: 0`, options that cannot hold together.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, type BookInput, type Calc, type CellInput } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { createCalcPool } from '../src/worker/index.js'
import { chainBook, setA1 } from './async-books.js'
import { inside, pool } from './fixtures/pool.js'

const open: Calc[] = []
const made = async (...a: Parameters<typeof pool>) => {
  const p = await pool(...a)
  open.push(p)
  return p
}
afterAll(() => Promise.all(open.map((p) => p.close())))

/** What `p` settles to, or 'hung' when it has not within `ms`. */
const settled = (p: Promise<unknown>, ms = 3000): Promise<unknown> =>
  Promise.race([p.then((v) => ['resolved', v], (e) => ['rejected', e]), new Promise((resolve) => setTimeout(() => resolve('hung'), ms))])

/** Only how `p` settled: 'resolved', 'rejected' or 'hung'. */
const outcome = async (p: Promise<unknown>): Promise<unknown> => {
  const r = await settled(p)
  return Array.isArray(r) ? r[0] : r
}

// a function in meta: no JSON value, and a structured clone cannot take it
const unsendable = { value: 1, meta: { f: () => 1 } } as unknown as CellInput
const bad: BookInput = { sheets: [{ name: 'S', cells: { A1: unsendable } }] }

describe('the pool at its edges', () => {
  it('a book or an edit that cannot cross to a thread is refused with InputError, and the thread goes on', async () => {
    const p = await made({ threads: 1 })
    await expect(p.compute(bad)).rejects.toThrow(InputError)
    await expect(p.compute(bad)).rejects.toThrow(/^book: cannot be sent to a thread/)
    expect(await outcome(p.compute(chainBook(2)))).toBe('resolved')
    // one running, the bad one waiting: the answer of the first sends the second, which fails alone
    const good = p.compute(synthBook({ seed: 1, rows: 340 }).book)
    const worse = p.compute(bad)
    await good
    expect(await settled(worse)).toMatchObject(['rejected', expect.any(InputError)])
    const s = await p.openSession(chainBook(2))
    await expect(s.setCells([{ sheet: 'S', cell: 'B9', input: unsendable }])).rejects.toThrow(/^edits: cannot be sent to a thread/)
    expect((await s.setCells(setA1(3))).changes.map((c) => c.cell)).toEqual(['A1', 'B1', 'B2'])
  }, 30_000)

  it('a session closed while the pool closes is closed, not left waiting', async () => {
    const p = await pool({ threads: 1 })
    const s = await p.openSession(chainBook(3))
    const busy = p.compute(synthBook({ seed: 1, rows: 1700 }).book).catch(() => undefined)
    const closing = s.close()
    await p.close()
    expect(await settled(closing)).toEqual(['resolved', undefined])
    await busy
  }, 30_000)

  it('a session closed while its thread is lost is closed', async () => {
    const { calc, crash } = await inside({ threads: 1 })
    open.push(calc)
    const s = await calc.openSession(chainBook(3))
    const closing = s.close()
    crash(0)
    expect(await settled(closing)).toEqual(['resolved', undefined])
    expect(await outcome(calc.compute(chainBook(2)))).toBe('resolved')
  }, 30_000)

  it('maxQueue 0: a task that an idle thread takes at once is no task waiting', async () => {
    const p = await made({ threads: 1, maxQueue: 0 })
    expect(await outcome(p.compute(chainBook(2)))).toBe('resolved')
    const running = p.compute(synthBook({ seed: 1, rows: 900 }).book)
    await expect(p.compute(chainBook(2))).rejects.toMatchObject({ code: 'queue_full' })
    await running
    const s = await p.openSession(chainBook(2))
    expect(await outcome(s.setCells(setA1(4)))).toBe('resolved')
  }, 30_000)

  it('maxHeapMb with a spawn of its own is refused: such a worker has no memory limit', () => {
    const worker = { postMessage() {}, terminate() {}, addEventListener() {} }
    expect(() => createCalcPool({ spawn: () => worker, maxHeapMb: 64 })).toThrow(InputError)
    expect(() => createCalcPool({ spawn: () => worker, maxHeapMb: 64 })).toThrow(/^opts\.maxHeapMb:/)
  })
})

