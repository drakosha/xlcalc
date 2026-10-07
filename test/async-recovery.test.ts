/**
 * The pool when a thread is lost and when it is full: a hard deadline, a thread's heap and a crash take the thread
 * down and the pool goes on, a session comes back on its next call from its snapshot and journal with the hash of a
 * session that never lost its thread, a session leaves a full thread and comes back the same way, a failed restore
 * closes that session only; admission (`maxQueue`, `maxSessions`) and the order in which a thread takes its tasks.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, compute, openSession, type BookInput, type Calc, type CellEdit } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { nodeSpawn, workerThreads } from '../src/worker/platform.js'
import { chainBook, CLOCK, setA1, slowBook } from './async-books.js'
import { inside, pool } from './fixtures/pool.js'
import { edits } from './session-edits.js'

const fails = async (p: Promise<unknown>): Promise<CalcError> => {
  try {
    await p
  } catch (e) {
    expect(e).toBeInstanceOf(CalcError)
    return e as CalcError
  }
  throw new Error('expected CalcError')
}

const open: Calc[] = []
afterAll(() => Promise.all(open.map((p) => p.close())))
const kept = <T extends Calc>(c: T) => (open.push(c), c)

const heavy: CellEdit[] = setA1(2)

describe('a thread lost', () => {
  it('past its hard deadline: the call fails, the session comes back with the hash of one that never lost its thread', async () => {
    const synth = synthBook({ seed: 1, rows: 340 })
    const book: BookInput = { ...synth.book, sheets: [...synth.book.sheets, ...slowBook().sheets] }
    const p = kept(await pool({ threads: 1, hardTimeoutMs: 400 }))
    const s = await p.openSession(book, { clock: CLOCK })
    const y = openSession(book, { clock: CLOCK })
    const list = edits(4, 12, synth)
    for (const batch of list.slice(0, 6)) expect(await s.setCells(batch)).toEqual(y.setCells(batch))
    const e = await fails(s.setCells(heavy))
    expect(e.code).toBe('hard_timeout')
    for (const batch of list.slice(6)) expect(await s.setCells(batch)).toEqual(y.setCells(batch))
    expect((await s.result()).hash).toBe(y.result().hash)
    // a one-off computation runs on the new thread
    expect((await p.compute(chainBook(3), { clock: CLOCK })).hash).toBe(compute(chainBook(3), { clock: CLOCK }).hash)
  }, 60_000)

  it("past the thread's heap (maxHeapMb): out_of_memory, and the process goes on", async () => {
    const p = kept(await pool({ threads: 1, maxHeapMb: 64, limits: { cells: 2_000_000 } }))
    const s = await p.openSession(chainBook(50), { clock: CLOCK })
    await s.setCells(setA1(5))
    // twenty arrays of a million values, each spilled down its column: far more than 64 MB, each well under it
    const cells: BookInput['sheets'][0]['cells'] = {}
    for (let c = 0; c < 20; c++) cells[String.fromCharCode(65 + c) + '1'] = { formula: 'SEQUENCE(1000000)' }
    expect((await fails(p.compute({ sheets: [{ name: 'S', cells }] }))).code).toBe('out_of_memory')
    expect((await s.result()).hash).toBe(compute({ sheets: [{ name: 'S', cells: { ...chainBook(50).sheets[0].cells, A1: { value: 5 } } }] }, { clock: CLOCK }).hash)
  }, 60_000)

  it('a crash: the waiting calls of every session of the thread go on; sessions come back only when called', async () => {
    const { calc, inside: x, crash } = await inside({ threads: 1 })
    kept(calc)
    const [a, b] = [chainBook(30), chainBook(40)]
    const [sa, sb] = [await calc.openSession(a, { clock: CLOCK }), await calc.openSession(b, { clock: CLOCK })]
    await sa.setCells(setA1(3))
    await sb.setCells(setA1(4))
    expect(x.slots[0].resident.size).toBe(2)
    crash(0)
    expect(x.slots[0].resident.size).toBe(0)
    expect((await sa.result()).hash).toBe(compute({ sheets: [{ name: 'S', cells: { ...a.sheets[0].cells, A1: { value: 3 } } }] }, { clock: CLOCK }).hash)
    expect(x.slots[0].resident.size).toBe(1)
    // a crash while a call runs: that call fails, the next one waiting is answered
    await sb.result()
    const p1 = sb.setCells(setA1(7))
    const p2 = sb.setCells(setA1(8))
    crash(0)
    expect((await fails(p1)).code).toBe('thread_failed')
    await p2
    expect((await sb.result()).hash).toBe(compute({ sheets: [{ name: 'S', cells: { ...b.sheets[0].cells, A1: { value: 8 } } }] }, { clock: CLOCK }).hash)
  }, 60_000)

  it('a session restored from a snapshot and a journal cut at it', async () => {
    const synth = synthBook({ seed: 1, rows: 200 })
    const { calc, inside: x, crash } = await inside({ threads: 1, snapshotEvery: 4 })
    kept(calc)
    const s = await calc.openSession(synth.book, { live: true })
    const list = edits(6, 10, synth)
    for (const batch of list) await s.setCells(batch)
    await s.result()
    const r = [...x.recs][0]
    expect(r.snapshot.seq).toBeGreaterThanOrEqual(8)
    expect(r.journal.length).toBeLessThan(4)
    const book = await s.book()
    const clock = s.clock
    crash(0)
    expect((await s.result()).hash).toBe(compute(book, { clock }).hash)
    expect(await s.book()).toEqual(book)
  }, 60_000)

  it('a restore that fails closes that session only', async () => {
    const { calc, inside: x, crash } = await inside({ threads: 1 })
    kept(calc)
    const [sa, sb] = [await calc.openSession(chainBook(5), { clock: CLOCK }), await calc.openSession(chainBook(6), { clock: CLOCK })]
    await sa.setCells(setA1(2))
    await sb.setCells(setA1(2))
    // the journal of a no longer replays: its edit names a sheet the book lacks
    const ra = [...x.recs].find((r) => r.journal.length && r.snapshot.book.sheets[0].cells.B5)!
    ra.journal[0].args = [[{ sheet: 'Gone', cell: 'A1', input: { value: 1 } }]]
    crash(0)
    const e = await fails(sa.setCells(setA1(3)))
    expect(e.code).toBe('restore_failed')
    expect((e.cause as Error).message).toMatch(/unknown sheet "Gone"/)
    expect((await fails(sa.result())).code).toBe('restore_failed')
    expect((await sb.setCells(setA1(9))).map((c) => c.cell)).toContain('B6')
  }, 60_000)

  it('a thread that exits by itself is reported as thread_failed', async () => {
    const wt = workerThreads()!
    const loss = await new Promise((resolve) => nodeSpawn(wt, new URL('data:text/javascript,process.exit(3)'), undefined)(() => {}, resolve))
    expect(loss).toEqual({ code: 'thread_failed', message: 'the thread exited with code 3' })
  })
})

describe('a full pool', () => {
  it('a thread over maxSessionsPerThread: the session called least recently leaves it and comes back on its next call', async () => {
    const { calc, inside: x } = await inside({ threads: 1, maxSessionsPerThread: 2 })
    kept(calc)
    const books = [chainBook(10), chainBook(20), chainBook(30)]
    const ss = [] as Awaited<ReturnType<Calc['openSession']>>[]
    for (const b of books) ss.push(await calc.openSession(b, { clock: CLOCK }))
    expect(x.slots[0].resident.size).toBe(2)
    const ys = books.map((b) => openSession(b, { clock: CLOCK }))
    for (let k = 0; k < 9; k++) {
      const i = (k * 2) % 3
      expect(await ss[i].setCells(setA1(k + 2))).toEqual(ys[i].setCells(setA1(k + 2)))
      expect(x.slots[0].resident.size).toBe(2)
    }
    for (const [i, s] of ss.entries()) expect((await s.result()).hash).toBe(ys[i].result().hash)
  }, 60_000)

  it('maxSessions refuses another session; maxQueue another waiting task', async () => {
    const p = kept(await pool({ threads: 1, maxSessions: 2, maxQueue: 1 }))
    const s1 = await p.openSession(chainBook(2))
    await p.openSession(chainBook(2))
    expect((await fails(p.openSession(chainBook(2)))).code).toBe('too_many_sessions')
    await s1.close()
    await p.openSession(chainBook(2))
    const big = synthBook({ seed: 1, rows: 1700 }).book
    const running = p.compute(big)
    const waiting = p.compute(chainBook(3))
    expect((await fails(p.compute(chainBook(3)))).code).toBe('queue_full')
    await Promise.all([running, waiting])
  }, 60_000)

  it("an idle thread takes its sessions' next call before a waiting one-off computation", async () => {
    const p = kept(await pool({ threads: 1 }))
    const s = await p.openSession(chainBook(3), { clock: CLOCK })
    const order: string[] = []
    const first = p.compute(synthBook({ seed: 1, rows: 900 }).book).then(() => order.push('first job'))
    const job = p.compute(chainBook(3)).then(() => order.push('job'))
    const call = s.setCells(setA1(5)).then(() => order.push('session'))
    await Promise.all([first, job, call])
    expect(order).toEqual(['first job', 'session', 'job'])
  }, 60_000)

  it('closing the pool refuses what waits and runs', async () => {
    const p = await pool({ threads: 1 })
    const s = await p.openSession(chainBook(3))
    const running = p.compute(synthBook({ seed: 1, rows: 1700 }).book)
    const waiting = s.setCells(setA1(2))
    await p.close()
    for (const q of [running, waiting, s.result(), p.compute(chainBook(2))]) expect((await fails(q)).code).toBe('closed')
  })
})
