/**
 * Undo and redo in the same thread and in a thread, with closures and eval: the answers, books and hashes of a
 * synchronous session, `mapCell` made on the calling side from the geometry a thread sends; an undo stopped by its
 * deadline or its signal changes nothing; a session whose thread was lost comes back with its history, from the journal
 * (undos and redos in it) and from a snapshot, and one that left a full thread too.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, createInlineCalc, createSession, type AsyncSession, type BookInput, type CellRef, type Execution, type Session, type UndoResult } from '../src/index.js'
import { CLOCK, heavyBook } from './async-books.js'
import { closedAfter, inside } from './fixtures/pool.js'

const book = (): BookInput => ({
  sheets: [
    { name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'SUM(A1:A9)' }, C1: { formula: 'SEQUENCE(A2)' } } },
    { name: 'R', cells: { A1: { formula: 'S!B1*2' } } },
  ],
})

/** Calls of every kind, each followed in turn by undos and redos. */
type Call = (s: Session | AsyncSession) => unknown
const CALLS: Call[] = [
  (s) => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }]),
  (s) => s.patchCells([{ sheet: 'S', cell: 'A2', set: { numberFormat: '0.00', note: { text: 'n' } } }]),
  (s) => s.editStructure({ op: 'insertRows', sheet: 'S', row: 2, count: 2 }),
  (s) => s.setLayout([{ sheet: 'S', merges: ['C2:D3'] }]),
  (s) => s.setNames([{ name: 'k', ref: '3' }]),
  (s) => s.setStyles([{ font: { bold: true } }]),
  (s) => s.editStructure({ op: 'renameSheet', sheet: 'R', name: 'Q' }),
  (s) => s.editStructure({ op: 'deleteSheet', sheet: 'S' }),
]
const UNDO = (s: Session | AsyncSession) => s.undo()
const REDO = (s: Session | AsyncSession) => s.redo()
const MOVES: Call[] = [...CALLS.slice(0, 4), UNDO, UNDO, REDO, ...CALLS.slice(4), UNDO, UNDO, UNDO, REDO, UNDO, UNDO, UNDO, REDO, REDO, (s) => s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 7 } }]), REDO, UNDO]

/** An answer without its mapCell, and where mapCell takes a few cells of the book before the call. */
function plain(a: unknown, sheets: string[]) {
  const r = a as (UndoResult & { book?: unknown }) | null
  if (!r || typeof r !== 'object' || !('mapCell' in r) || !r.mapCell) return a
  const { mapCell, book: _, ...rest } = r
  const at = sheets.flatMap((sheet) => ['A1', 'A2', 'A4', 'C3'].map((cell): CellRef => ({ sheet, cell })))
  return { ...rest, at: at.map((ref) => mapCell(ref)) }
}

const { keep, made } = closedAfter()

describe.each(['closures', 'eval'] as const)('undo and redo without blocking, %s', (execution: Execution) => {
  const o = { clock: CLOCK, execution }

  /** The moves on `s` give what they give on a synchronous session, with its book, hash and history after each. */
  async function same(s: AsyncSession, moves = MOVES, y = createSession(book(), o)) {
    for (const [k, move] of moves.entries()) {
      const sheets = y.getBook().sheets.map((x) => x.name)
      const want = plain(move(y), sheets)
      expect(plain(await move(s), sheets), `move ${k}`).toEqual(want)
      expect(s.history, `move ${k}`).toEqual(y.history)
      expect(await s.getBook(), `move ${k}`).toEqual(y.getBook())
      expect(await s.getHash(), `move ${k}`).toBe(y.getHash())
    }
    return y
  }

  it('in the same thread', async () => {
    const calc = keep(createInlineCalc({ sliceMs: 1 }))
    await same(await calc.createSession(book(), o))
  })

  it('in a thread: mapCell from the geometry the thread sends', async () => {
    const calc = await made({ threads: 1 })
    await same(await calc.createSession(book(), o))
  }, 60_000)

  it('an undo stopped by its deadline or its signal changes nothing', async () => {
    const calc = keep(createInlineCalc({ sliceMs: 1 }))
    const s = await calc.createSession(heavyBook(), o)
    await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }])
    const [b, hash] = [JSON.stringify(await s.getBook()), await s.getHash()]
    const e = await s.undo({ timeoutMs: 1 }).catch((x) => x)
    expect(e).toBeInstanceOf(CalcError)
    expect(e.code).toBe('soft_timeout')
    const stop = new AbortController()
    const waiting = s.undo({ signal: stop.signal }).catch((x) => x)
    stop.abort()
    expect((await waiting).code).toBe('aborted')
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    expect(JSON.stringify(await s.getBook())).toBe(b)
    expect(await s.getHash()).toBe(hash)
    // B1 of the heavy book is too large to change by 1
    expect((await s.undo())!.changes).toEqual([{ sheet: 'S', cell: 'A1', value: 1, text: '1' }])
    expect(s.history).toEqual({ undo: 0, redo: 1 })
  })

  it('a thread lost: the session comes back with its history, from the journal and from a snapshot taken after undos', async () => {
    for (const snapshotEvery of [64, 2]) {
      const { calc, inside: x, crash } = await inside({ threads: 1, snapshotEvery })
      keep(calc)
      const s = await calc.createSession(book(), o)
      const y = await same(s, MOVES.slice(0, 13))
      const r = [...x.recs][0]
      if (snapshotEvery === 64) expect(r.journal.map((e) => e.op)).toContain('undo')
      // the snapshot holds the history as text
      else expect(JSON.parse((r.snapshot.history as { json: string }).json).undo.length).toBeGreaterThan(0)
      crash(0)
      await same(s, MOVES.slice(13), y)
    }
  }, 120_000)

  it("a session in a pool with maxHeapMb keeps its share of the heap by default: 8 MB of 384 among 6 sessions; its own maxMb wins", async () => {
    // 20,000 rows of 10 values: a step that deletes 1,000 of them keeps about 1.2 MB, one of 4,000 about 5 MB
    const cells: BookInput['sheets'][number]['cells'] = {}
    for (let r = 1; r <= 20_000; r++) for (let c = 0; c < 10; c++) cells[`${'ABCDEFGHIJ'[c]}${r}`] = { value: r * 10 + c }
    const big = (): BookInput => ({ sheets: [{ name: 'S', cells: { ...cells } }] })
    const { calc, inside: x } = await inside({ threads: 1, maxHeapMb: 384, maxSessionsPerThread: 6, snapshotEvery: 1 })
    keep(calc)
    const [s, own] = [await calc.createSession(big(), o), await calc.createSession(big(), { ...o, undo: { maxMb: 64 } })]
    const y = createSession(big(), { ...o, undo: { maxMb: 8 } })
    const counts: number[] = []
    for (const count of [1000, 1000, 1000, 1000, 4000]) {
      const op = { op: 'deleteRows', sheet: 'S', row: 1, count } as const
      await s.editStructure(op)
      await own.editStructure(op)
      y.editStructure(op)
      expect(s.history).toEqual(y.history)
      counts.push(s.history.undo)
      // what the snapshot taken after the call holds: within 8 MB by the estimate
      const r = [...x.recs].find((rec) => rec.sid === Math.min(...[...x.recs].map((z) => z.sid)))!
      const kept = JSON.parse((r.snapshot.history as { json: string }).json) as { undo: { size: number }[] }
      expect(kept.undo.reduce((n, step) => n + step.size, 0)).toBeLessThanOrEqual(8 * 2 ** 20)
    }
    // the oldest steps went when a large one came
    expect(counts).toEqual([1, 2, 3, 4, 3])
    expect(own.history).toEqual({ undo: 5, redo: 0 })
    for (let k = 0; k < 3; k++) {
      expect((await s.undo())!.call).toBe('editStructure')
      y.undo()
      expect(await s.getHash()).toBe(y.getHash())
    }
    expect(await s.undo()).toBeNull()
    expect(await s.getBook()).toEqual(y.getBook())
    // the book after the first two deletes, of 1,000 rows each
    expect(Object.keys((await s.getBook()).sheets[0].cells)).toHaveLength(180_000)
  }, 120_000)

  it('sessions that take turns on a full thread keep their histories', async () => {
    const calc = await made({ threads: 1, maxSessionsPerThread: 1 })
    const [a, b] = [await calc.createSession(book(), o), await calc.createSession(book(), o)]
    const [ya, yb] = [createSession(book(), o), createSession(book(), o)]
    for (const [k, move] of MOVES.entries()) {
      for (const [s, y] of [[a, ya], [b, yb]] as const) {
        const sheets = y.getBook().sheets.map((x) => x.name)
        expect(plain(await move(s), sheets), `move ${k}`).toEqual(plain(move(y), sheets))
        expect(s.history).toEqual(y.history)
      }
    }
    expect(await a.getBook()).toEqual(ya.getBook())
    expect(await b.getHash()).toBe(yb.getHash())
  }, 120_000)
})
