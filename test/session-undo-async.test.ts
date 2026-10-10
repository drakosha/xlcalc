/**
 * Undo and redo in the same thread and in a thread, with closures and eval: the answers, books and hashes of a
 * synchronous session, `mapCell` made on the calling side from the geometry a thread sends; an undo stopped by its
 * deadline or its signal changes nothing; a session whose thread was lost comes back with its history, from the journal
 * (undos and redos in it) and from a snapshot, and one that left a full thread too.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, createInlineCalc, openSession, type AsyncSession, type BookInput, type Calc, type CellRef, type Execution, type Session, type UndoResult } from '../src/index.js'
import { CLOCK, heavyBook } from './async-books.js'
import { inside, pool } from './fixtures/pool.js'

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
  (s) => s.patchCells([{ sheet: 'S', cell: 'A2', set: { format: '0.00', note: { text: 'n' } } }]),
  (s) => s.restructure({ op: 'insertRows', sheet: 'S', row: 2, count: 2 }),
  (s) => s.setLayout([{ sheet: 'S', merges: ['C2:D3'] }]),
  (s) => s.setNames([{ name: 'k', ref: '3' }]),
  (s) => s.setStyles([{ font: { bold: true } }]),
  (s) => s.restructure({ op: 'renameSheet', sheet: 'R', name: 'Q' }),
  (s) => s.restructure({ op: 'deleteSheet', sheet: 'S' }),
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

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

describe.each(['closures', 'eval'] as const)('undo and redo without blocking, %s', (execution: Execution) => {
  const o = { clock: CLOCK, execution }

  /** The moves on `s` give what they give on a synchronous session, with its book, hash and history after each. */
  async function same(s: AsyncSession, moves = MOVES, y = openSession(book(), o)) {
    for (const [k, move] of moves.entries()) {
      const sheets = y.book().sheets.map((x) => x.name)
      const want = plain(move(y), sheets)
      expect(plain(await move(s), sheets), `move ${k}`).toEqual(want)
      expect(s.history, `move ${k}`).toEqual(y.history)
      expect(await s.book(), `move ${k}`).toEqual(y.book())
      expect(await s.hash(), `move ${k}`).toBe(y.hash())
    }
    return y
  }

  it('in the same thread', async () => {
    const calc = createInlineCalc({ sliceMs: 1 })
    calcs.push(calc)
    await same(await calc.openSession(book(), o))
  })

  it('in a thread: mapCell from the geometry the thread sends', async () => {
    const calc = await pool({ threads: 1 })
    calcs.push(calc)
    await same(await calc.openSession(book(), o))
  }, 60_000)

  it('an undo stopped by its deadline or its signal changes nothing', async () => {
    const calc = createInlineCalc({ sliceMs: 1 })
    calcs.push(calc)
    const s = await calc.openSession(heavyBook(), o)
    await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }])
    const [b, hash] = [JSON.stringify(await s.book()), await s.hash()]
    const e = await s.undo({ timeoutMs: 1 }).catch((x) => x)
    expect(e).toBeInstanceOf(CalcError)
    expect(e.code).toBe('soft_timeout')
    const stop = new AbortController()
    const waiting = s.undo({ signal: stop.signal }).catch((x) => x)
    stop.abort()
    expect((await waiting).code).toBe('aborted')
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    expect(JSON.stringify(await s.book())).toBe(b)
    expect(await s.hash()).toBe(hash)
    // B1 of the heavy book is too large to change by 1
    expect((await s.undo())!.changes).toEqual([{ sheet: 'S', cell: 'A1', value: 1, text: '1' }])
    expect(s.history).toEqual({ undo: 0, redo: 1 })
  })

  it('a thread lost: the session comes back with its history, from the journal and from a snapshot taken after undos', async () => {
    for (const snapshotEvery of [64, 2]) {
      const { calc, inside: x, crash } = await inside({ threads: 1, snapshotEvery })
      calcs.push(calc)
      const s = await calc.openSession(book(), o)
      const y = await same(s, MOVES.slice(0, 13))
      const r = [...x.recs][0]
      if (snapshotEvery === 64) expect(r.journal.map((e) => e.op)).toContain('undo')
      // the snapshot holds the history as text
      else expect(JSON.parse((r.snapshot.history as { json: string }).json).undo.length).toBeGreaterThan(0)
      crash(0)
      await same(s, MOVES.slice(13), y)
    }
  }, 120_000)

  it('sessions that take turns on a full thread keep their histories', async () => {
    const calc = await pool({ threads: 1, maxSessionsPerThread: 1 })
    calcs.push(calc)
    const [a, b] = [await calc.openSession(book(), o), await calc.openSession(book(), o)]
    const [ya, yb] = [openSession(book(), o), openSession(book(), o)]
    for (const [k, move] of MOVES.entries()) {
      for (const [s, y] of [[a, ya], [b, yb]] as const) {
        const sheets = y.book().sheets.map((x) => x.name)
        expect(plain(await move(s), sheets), `move ${k}`).toEqual(plain(move(y), sheets))
        expect(s.history).toEqual(y.history)
      }
    }
    expect(await a.book()).toEqual(ya.book())
    expect(await b.hash()).toBe(yb.hash())
  }, 120_000)
})
