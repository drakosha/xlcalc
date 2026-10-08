/**
 * Computation in the same thread, in portions (`createInlineCalc`): the same values and hashes as the synchronous
 * engine in both execution modes, computing in slices that let timers run, a soft deadline and an abort as
 * `CalcError`, calls taken in order with their own changes, a stopped call that changes nothing, limits, closing.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, compute, createInlineCalc, openSession, type BookInput } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { BOOKS, chainBook, CLOCK, heavyBook, setA1 } from './async-books.js'
import { apply, edits } from './session-edits.js'

const fails = async (p: Promise<unknown>): Promise<CalcError> => {
  try {
    await p
  } catch (e) {
    expect(e).toBeInstanceOf(CalcError)
    return e as CalcError
  }
  throw new Error('expected CalcError')
}

/** Runs of the event loop while `p` is pending. */
async function ticksDuring(p: Promise<unknown>): Promise<number> {
  let ticks = 0
  let done = false
  const tick = () => {
    ticks++
    if (!done) setImmediate(tick)
  }
  setImmediate(tick)
  await p.finally(() => (done = true))
  return ticks
}

describe.each(['closures', 'eval'] as const)('in the same thread, execution %s', (execution) => {
  it('the case books: values and hashes as the synchronous engine', async () => {
    const calc = createInlineCalc({ sliceMs: 1 })
    for (const [name, book] of BOOKS) {
      const sync = compute(book, { clock: CLOCK, execution })
      const r = await calc.compute(book, { clock: CLOCK, execution })
      expect(r.hash, name).toBe(sync.hash)
      expect(r.sheets, name).toEqual(sync.sheets)
      expect(r.issues, name).toEqual(sync.issues)
    }
  }, 120_000)

  it('a session over 40 edits: the same changes and hash after each as a synchronous session', async () => {
    const synth = synthBook({ seed: 1, rows: 340 })
    const calc = createInlineCalc({ sliceMs: 0 })
    const s = await calc.openSession(synth.book, { clock: CLOCK, execution })
    const sync = openSession(synth.book, { clock: CLOCK, execution })
    for (const batch of edits(5, 40, synth)) {
      expect(await s.setCells(batch)).toEqual(sync.setCells(batch))
      expect((await s.result()).hash).toBe(sync.result().hash)
    }
    await s.close()
  }, 120_000)
})

describe('in the same thread', () => {
  it('computes in slices: the event loop runs while a book computes', async () => {
    const calc = createInlineCalc({ sliceMs: 2 })
    const book = synthBook({ seed: 1, rows: 1700 }).book
    const p = calc.compute(book, { clock: CLOCK })
    expect(await ticksDuring(p)).toBeGreaterThan(5)
    expect((await p).hash).toBe(compute(book, { clock: CLOCK }).hash)
  }, 60_000)

  it('a soft deadline and an abort stop the computation with CalcError', async () => {
    const calc = createInlineCalc({ sliceMs: 1 })
    const book = synthBook({ seed: 1, rows: 1700 }).book
    expect((await fails(calc.compute(book, { clock: CLOCK, timeoutMs: 5 }))).code).toBe('soft_timeout')
    const ac = new AbortController()
    const p = calc.compute(book, { clock: CLOCK, signal: ac.signal })
    setImmediate(() => ac.abort('enough'))
    const e = await fails(p)
    expect([e.code, e.cause]).toEqual(['aborted', 'enough'])
    expect((await fails(calc.compute(book, { signal: AbortSignal.abort() }))).code).toBe('aborted')
    // the calculator's default deadline applies to every call
    expect((await fails(createInlineCalc({ sliceMs: 1, timeoutMs: 5 }).compute(book, { clock: CLOCK }))).code).toBe('soft_timeout')
  }, 60_000)

  it('formulas of any cost: the slice and the soft deadline hold', async () => {
    // 300 formulas that each sort 10,000 values: a few milliseconds each, where a check counted formulas
    const cells: BookInput['sheets'][0]['cells'] = {}
    for (let r = 1; r <= 10_000; r++) cells[`A${r}`] = { value: (r * 7919) % 10_007 }
    for (let r = 1; r <= 300; r++) cells[`B${r}`] = { formula: 'INDEX(SORT(A$1:A$10000),ROW())' }
    const book = { sheets: [{ name: 'S', cells }] }
    const calc = createInlineCalc({ sliceMs: 10 })
    let last = performance.now()
    let worst = 0
    let done = false
    const tick = () => {
      const now = performance.now()
      worst = Math.max(worst, now - last)
      last = now
      if (!done) setImmediate(tick)
    }
    setImmediate(tick)
    const t0 = performance.now()
    const e = await fails(calc.compute(book, { timeoutMs: 100 }))
    const took = performance.now() - t0
    done = true
    expect(e.code).toBe('soft_timeout')
    expect(took).toBeLessThan(200)
    // a slice and the formula in flight when it ends (a sort, a few ms; more while the code is cold), with room for a loaded machine
    expect(worst).toBeLessThan(50)
  }, 60_000)

  it('the call after a stopped one: computing the book again does not count in its soft deadline', async () => {
    const calc = createInlineCalc({ sliceMs: 5 })
    // a cold computation of about 150 ms, longer than the calls' deadline of 50 ms
    const s = await calc.openSession(heavyBook(), { clock: CLOCK })
    expect((await fails(s.setCells(setA1(2), { timeoutMs: 50 }))).code).toBe('soft_timeout')
    expect(await s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 5 } }], { timeoutMs: 50 })).toEqual([{ sheet: 'S', cell: 'C1', value: 5, text: '5' }])
    expect((await s.result({ timeoutMs: 50 })).hash).toBe(compute({ sheets: [{ name: 'S', cells: { ...heavyBook().sheets[0].cells, C1: { value: 5 } } }] }, { clock: CLOCK }).hash)
    // a signal still stops it, between the formulas of the book
    const t = await calc.openSession(chainBook(20_000), { clock: CLOCK })
    expect((await fails(t.setCells(setA1(2), { signal: AbortSignal.abort() }))).code).toBe('aborted')
    const stop = new AbortController()
    const first = t.setCells(setA1(3), { signal: stop.signal })
    setImmediate(() => stop.abort())
    expect((await fails(first)).code).toBe('aborted')
    const again = new AbortController()
    const next = t.result({ signal: again.signal })
    setImmediate(() => again.abort())
    expect((await fails(next)).code).toBe('aborted')
  }, 60_000)

  it('calls made without waiting are taken in order, each with its own changes', async () => {
    const book = chainBook(300)
    const calc = createInlineCalc({ sliceMs: 0 })
    const s = await calc.openSession(book, { clock: CLOCK, display: false })
    const sync = openSession(book, { clock: CLOCK, display: false })
    const values = [5, 6, 7, 8, 9, 10]
    const all = await Promise.all(values.map((v) => s.setCells(setA1(v))))
    for (const [i, v] of values.entries()) expect(all[i]).toEqual(sync.setCells(setA1(v)))
    expect((await s.result()).hash).toBe(sync.result().hash)
  })

  it('two sessions with their calls mixed, without waiting: each has its own changes and hash', async () => {
    const calc = createInlineCalc({ sliceMs: 0 })
    const [a, b] = [chainBook(400), synthBook({ seed: 2, rows: 200 }).book]
    const [sa, sb] = await Promise.all([calc.openSession(a, { clock: CLOCK }), calc.openSession(b, { clock: CLOCK })])
    const [ya, yb] = [openSession(a, { clock: CLOCK }), openSession(b, { clock: CLOCK })]
    const eb = edits(9, 6, synthBook({ seed: 2, rows: 200 }))
    const pending: Promise<unknown>[] = []
    const expected: unknown[] = []
    for (let i = 0; i < 6; i++) {
      pending.push(sa.setCells(setA1(i + 2)), sb.setCells(eb[i]))
      expected.push(ya.setCells(setA1(i + 2)), yb.setCells(eb[i]))
    }
    expect(await Promise.all(pending)).toEqual(expected)
    expect([(await sa.result()).hash, (await sb.result()).hash]).toEqual([ya.result().hash, yb.result().hash])
  })

  it('a call stopped in the middle changes nothing; the session goes on as if it had not been made', async () => {
    const book = chainBook(5000)
    const calc = createInlineCalc({ sliceMs: 0 })
    const s = await calc.openSession(book, { clock: CLOCK })
    const before = await s.result()
    // stopped at its first yield
    const ac = new AbortController()
    const p = s.setCells(setA1(42), { signal: ac.signal })
    setImmediate(() => ac.abort())
    expect((await fails(p)).code).toBe('aborted')
    expect((await s.book()).sheets[0].cells.A1).toEqual({ value: 1 })
    expect((await s.result()).hash).toBe(before.hash)
    // a deadline too; a layout and a clock call put back what they changed
    expect((await fails(s.setCells(setA1(43), { timeoutMs: 0 }))).code).toBe('soft_timeout')
    expect((await fails(s.setLayout([{ sheet: 'S', rows: '3', set: { hidden: true } }, { sheet: 'S', columns: 'B', set: { width: 3 } }], { signal: AbortSignal.abort() }))).code).toBe('aborted')
    expect((await s.book()).sheets[0].columns).toBeUndefined()
    const sync = openSession(book, { clock: CLOCK })
    expect(await s.setCells(setA1(7))).toEqual(sync.setCells(setA1(7)))
    expect((await s.result()).hash).toBe(sync.result().hash)
    expect((await s.result()).hash).toBe(compute(apply(book, setA1(7)), { clock: CLOCK }).hash)
  }, 60_000)

  it('a stopped setNames or restructure keeps the old state; restructure answers mapCell and changes', async () => {
    const book = chainBook(3000)
    const calc = createInlineCalc({ sliceMs: 0 })
    const s = await calc.openSession(book, { live: true })
    const clock = s.clock
    const ac = new AbortController()
    const p = s.restructure({ op: 'insertRows', sheet: 'S', row: 1 }, { signal: ac.signal })
    setImmediate(() => ac.abort())
    expect((await fails(p)).code).toBe('aborted')
    expect(s.clock).toEqual(clock)
    expect((await s.book()).sheets[0].cells.A1).toEqual({ value: 1 })
    const r = await s.restructure({ op: 'insertRows', sheet: 'S', row: 1 })
    if (!r.ok) throw new Error(r.message)
    expect(r.mapCell({ sheet: 'S', cell: 'A1' })).toEqual({ sheet: 'S', cell: 'A2' })
    expect('book' in r).toBe(false)
    expect((await s.result()).hash).toBe(compute(await s.book(), { clock: s.clock }).hash)
  }, 60_000)

  it('limits give CalcError, from the calculator or from the call', async () => {
    const calc = createInlineCalc({ limits: { cells: 1000 } })
    const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'SUM(SEQUENCE(100,100))' } } }] }
    const e = await fails(calc.compute(book))
    expect([e.code, e.limit, e.cell]).toEqual(['limit', 'cells', { sheet: 'S', cell: 'A1' }])
    expect((await calc.compute(book, { limits: { cells: 10_000 } })).sheets[0].cells.A1.value).toBe(50_005_000)
    const s = await calc.openSession({ sheets: [{ name: 'S', cells: {} }] })
    expect((await fails(s.setCells([{ sheet: 'S', cell: 'A1', input: { formula: 'SUM(SEQUENCE(100,100))' } }]))).limit).toBe('cells')
  })

  it('closing: the session refuses calls after close; the calculator stops what runs and waits', async () => {
    const calc = createInlineCalc({ sliceMs: 0 })
    const s = await calc.openSession(chainBook(100), { clock: CLOCK })
    const last = s.setCells(setA1(2))
    const closed = s.close()
    expect((await fails(s.setCells(setA1(3)))).code).toBe('closed')
    await last
    await closed
    const t = await calc.openSession(chainBook(3000), { clock: CLOCK })
    const running = t.setCells(setA1(5))
    const waiting = t.setCells(setA1(6))
    const job = calc.compute(synthBook({ seed: 1, rows: 1700 }).book)
    await calc.close()
    for (const p of [running, waiting, job]) expect((await fails(p)).code).toBe('closed')
    expect((await fails(calc.compute(chainBook(2)))).code).toBe('closed')
  }, 60_000)

  it('a session closed before its calculator closes is closed; its running call is stopped', async () => {
    const calc = createInlineCalc({ sliceMs: 0 })
    const s = await calc.openSession(chainBook(3000), { clock: CLOCK })
    const running = s.setCells(setA1(2)).catch((e) => e)
    const closing = s.close().then(() => 'resolved', (e) => e)
    await calc.close()
    expect(await closing).toBe('resolved')
    expect((await running).code).toBe('closed')
  })

  it('options are checked as input', async () => {
    expect(() => createInlineCalc({ sliceMs: -1 })).toThrow('opts.sliceMs: expected a number of at least 0')
    await expect(createInlineCalc().compute(chainBook(2), { timeoutMs: 'x' } as never)).rejects.toThrow('opts.timeoutMs')
    await expect(createInlineCalc().compute(chainBook(2), { signal: {} } as never)).rejects.toThrow('opts.signal: expected an AbortSignal')
    const s = await createInlineCalc().openSession(chainBook(2))
    await expect(s.setCells(setA1(1), { deadline: 1 } as never)).rejects.toThrow('opts: unknown field "deadline"')
  })
})
