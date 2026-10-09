/**
 * Long steps in the same thread (`createInlineCalc`): a large book's result, hash and cells give way to the event loop
 * between portions, and a structure operation stops at its soft deadline, leaving the session as it was.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, createInlineCalc, openSession, restructure, type BookInput, type CellInput, type StructureOp } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { CLOCK, medianMs } from './async-books.js'

const SLICE_MS = 5

const fails = async (p: Promise<unknown>): Promise<CalcError> => {
  try {
    await p
  } catch (e) {
    expect(e).toBeInstanceOf(CalcError)
    return e as CalcError
  }
  throw new Error('expected CalcError')
}

/** The longest wait of the event loop while `p` is pending, and its runs; a call that never yields leaves it none. */
async function loopDuring<T>(p: Promise<T>): Promise<{ value: T; worst: number; ticks: number }> {
  let last = performance.now()
  let worst = 0
  let ticks = 0
  let done = false
  const tick = () => {
    const now = performance.now()
    worst = Math.max(worst, now - last)
    last = now
    ticks++
    if (!done) setImmediate(tick)
  }
  setImmediate(tick)
  const value = await p.finally(() => (done = true))
  return { value, worst: Math.max(worst, performance.now() - last), ticks }
}

/** `n` formulas over the table T: renaming it rewrites each of them. */
function tableBook(n: number): BookInput {
  const cells: Record<string, CellInput> = { A1: { value: 'x' }, B1: { value: 'y' }, A2: { value: 1 }, B2: { value: 2 } }
  for (let r = 1; r <= n; r++) cells[`D${r}`] = { formula: `SUM(T[x])+${r}` }
  return { sheets: [{ name: 'S', cells, tables: [{ name: 'T', range: 'A1:B2', columns: [{ name: 'x' }, { name: 'y' }] }] }] }
}

describe('long steps in the same thread', () => {
  it("a large book's result, hash and cells give way to the event loop, with the synchronous values", async () => {
    const book = synthBook({ seed: 1, rows: 8500 }).book
    const sync = openSession(book, { clock: CLOCK })
    const calc = createInlineCalc({ sliceMs: SLICE_MS })
    const s = await calc.openSession(book, { clock: CLOCK })
    const calls: [string, () => unknown, () => Promise<unknown>, (v: unknown) => unknown][] = [
      ['result', () => sync.result(), () => s.result(), (v) => JSON.stringify({ ...(v as object), ms: 0 })],
      ['hash', () => sync.hash(), () => s.hash(), (v) => v],
      ['cells', () => sync.cells('Данные', 'A1:Z99999'), () => s.cells('Данные', 'A1:Z99999'), (v) => JSON.stringify(v)],
    ]
    for (const [name, once, call, shown] of calls) {
      // the call's time in one stretch, as fast as the machine runs it now
      const whole = await medianMs(3, once)
      const { value, worst, ticks } = await loopDuring(call())
      expect(shown(value), name).toBe(shown(once()))
      // without yields the loop waited for the whole call and never ran
      expect(ticks, name).toBeGreaterThan(5)
      expect(worst, name).toBeLessThan(SLICE_MS + whole / 4)
    }
    await calc.close()
  }, 120_000)

  it('a structure operation stops at its soft deadline; the session is as it was and goes on', async () => {
    const synth = synthBook({ seed: 1, rows: 8500 }).book
    const ops: [BookInput, StructureOp][] = [
      [synth, { op: 'insertRows', sheet: 'Данные', row: 2 }],
      [synth, { op: 'deleteColumns', sheet: 'Данные', column: 'B' }],
      [synth, { op: 'renameSheet', sheet: 'Данные', name: 'D' }],
      [synth, { op: 'moveRange', from: { sheet: 'Данные', range: 'A2:C100' }, to: { sheet: 'Данные', cell: 'AA2' } }],
      [tableBook(60_000), { op: 'renameTable', table: 'T', name: 'U' }],
    ]
    const calc = createInlineCalc({ sliceMs: SLICE_MS })
    const timeoutMs = 20
    for (const [book, op] of ops) {
      const s = await calc.openSession(book, { clock: CLOCK })
      const before = await s.hash()
      // the operation's time without a deadline, as fast as the machine runs it now
      const whole = await medianMs(3, () => restructure(book, op))
      const t0 = performance.now()
      const e = await fails(s.restructure(op, { timeoutMs }))
      const took = performance.now() - t0
      expect(e.code, op.op).toBe('soft_timeout')
      // without checks inside the operation the deadline was seen after it, when the new book was computed
      expect(took, op.op).toBeLessThan(timeoutMs + whole / 4)
      expect(await s.hash(), op.op).toBe(before)
      const sync = openSession(book, { clock: CLOCK })
      const r = await s.restructure(op)
      const y = sync.restructure(op)
      if (!r.ok || !y.ok) throw new Error(`${op.op} refused`)
      expect(r.changes, op.op).toEqual(y.changes)
      expect(await s.hash(), op.op).toBe(sync.hash())
      await s.close()
    }
    await calc.close()
  }, 120_000)
})
