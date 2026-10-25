/**
 * Seeded random sessions that undo and redo (`undo-fuzz.ts`): changes of every kind, undos, redos and new changes
 * after them, against the books the changes made one after another. After each move the session's book equals the
 * book of its place in the history, its hash that of a full computation of that book, its history counts the steps
 * each way, and a screen that moves its values by `mapCell` and applies the changes of every call holds the session's
 * values; a refused change leaves the book and the history as they were. In a synchronous session, in the same thread
 * and in a thread, with closures and eval.
 */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, InputError, createSession, type BookInput, type CellChange, type CellRef, type ComputeResult, type Execution, type SessionHistory, type SessionOptions, type UndoResult } from '../src/index.js'
import { closedAfter, inside } from './fixtures/pool.js'
import { undoBook, undoMoves, type Change } from './undo-fuzz.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** What the test needs of a session, synchronous or not. */
export interface Undoable {
  setCells: (...a: never[]) => unknown
  history: SessionHistory
  undo(): Promise<UndoResult | null> | UndoResult | null
  redo(): Promise<UndoResult | null> | UndoResult | null
  getBook(): Promise<BookInput> | BookInput
  getHash(): Promise<string> | string
  getResult(): Promise<ComputeResult> | ComputeResult
  close(): unknown
}

type Answer = { ok?: boolean; changes?: CellChange[]; mapCell?: (ref: CellRef) => CellRef | null } | CellChange[] | void

/** A change through the session's call of its kind. */
function call(s: Undoable, c: Change): Promise<Answer> | Answer {
  const x = s as unknown as Record<string, (a: unknown) => Promise<Answer> | Answer>
  if ('edits' in c) return x.setCells(c.edits)
  if ('op' in c) return x.editStructure(c.op)
  if ('patches' in c) return x.patchCells(c.patches)
  if ('names' in c) return x.setNames(c.names)
  if ('styles' in c) return x.setStyles(c.styles)
  return x.setLayout(c.layout)
}

/** A screen: every cell's value by `sheet!cell`, as a host keeps them. */
type Screen = Map<string, string>

function screenOf(r: ComputeResult): Screen {
  const out: Screen = new Map()
  for (const s of r.sheets) for (const [a, c] of Object.entries(s.cells)) if (c.value !== null) out.set(`${s.name}!${a}`, JSON.stringify(c.value))
  return out
}

/** The screen after a call: its values moved by `mapCell`, then the changes applied. */
function moved(screen: Screen, changes: CellChange[], mapCell?: (ref: CellRef) => CellRef | null): Screen {
  let out = screen
  if (mapCell) {
    out = new Map()
    for (const [k, v] of screen) {
      const i = k.lastIndexOf('!')
      const to = mapCell({ sheet: k.slice(0, i), cell: k.slice(i + 1) })
      if (to) out.set(`${to.sheet}!${to.cell}`, v)
    }
  }
  for (const c of changes) {
    if (c.value === null) out.delete(`${c.sheet}!${c.cell}`)
    else out.set(`${c.sheet}!${c.cell}`, JSON.stringify(c.value))
  }
  return out
}

interface Saved {
  book: BookInput
  hash: string
}

/**
 * `moves` moves of seed `seed` on a session `open` gives; `between` runs after each (a thread lost, say). The history
 * keeps `depth` steps.
 */
export async function undoRun(seed: number, moves: number, open: (book: BookInput, o: SessionOptions) => Promise<Undoable>, o: SessionOptions & { clock: typeof clock }, between?: (k: number) => void) {
  const depth = o.undo?.depth ?? 100
  const book = undoBook(seed)
  const saved = (b: BookInput): Saved => ({ book: structuredClone(b), hash: compute(b, { clock: o.clock, execution: o.execution, display: false }).hash })
  const done: Saved[] = [saved(book)]
  const undone: Saved[] = []
  const s = await open(book, o)
  const next = undoMoves(seed)
  let screen = screenOf(await s.getResult())
  let steps = 0
  try {
    for (let k = 0; k < moves; k++) {
      between?.(k)
      const m = next()
      const at = `seed ${seed}, move ${k}: ${JSON.stringify(m)}`
      if ('change' in m) {
        const [before, history] = [JSON.stringify(await s.getBook()), s.history]
        let r: Answer
        try {
          r = await call(s, m.change)
        } catch (e) {
          if (!(e instanceof InputError)) throw e
          r = { ok: false }
        }
        if (r && 'ok' in r && !r.ok) {
          // refused: nothing changes
          expect(JSON.stringify(await s.getBook()), at).toBe(before)
          expect(s.history, at).toEqual(history)
          continue
        }
        steps++
        done.push(saved(await s.getBook()))
        if (done.length > depth + 1) done.shift()
        undone.length = 0
        screen = Array.isArray(r) ? moved(screen, r) : r ? moved(screen, r.changes!, r.mapCell) : screen
      } else {
        const back = 'undo' in m
        for (let i = 0; i < (back ? m.undo : m.redo); i++) {
          const r = await (back ? s.undo() : s.redo())
          const [from, to] = back ? [done, undone] : [undone, done]
          if (from.length === (back ? 1 : 0)) {
            expect(r, at).toBeNull()
            continue
          }
          expect(r, at).not.toBeNull()
          to.push(from.pop()!)
          screen = moved(screen, r!.changes, r!.mapCell)
        }
      }
      const want = done[done.length - 1]
      expect(s.history, at).toEqual({ undo: done.length - 1, redo: undone.length })
      expect(await s.getBook(), at).toEqual(want.book)
      expect(await s.getHash(), at).toBe(want.hash)
      expect(screen, at).toEqual(screenOf(await s.getResult()))
    }
  } finally {
    await s.close()
  }
  return steps
}

const sync = async (book: BookInput, o: SessionOptions) => createSession(book, o) as unknown as Undoable

const { keep, made } = closedAfter()

describe.each(['closures', 'eval'] as const)('undo and redo against the books of the changes, %s', (execution: Execution) => {
  it.each([1, 2, 3, 4, 5, 6])('seed %i, synchronous', async (seed) => {
    expect(await undoRun(seed, 90, sync, { clock, execution, display: seed % 2 === 0 })).toBeGreaterThan(40)
  })

  it('seed 7, a history of 5 steps', async () => {
    await undoRun(7, 90, sync, { clock, execution, display: false, undo: { depth: 5 } })
  })

  it('seed 8, in the same thread', async () => {
    const calc = keep(createInlineCalc({ sliceMs: 1 }))
    await undoRun(8, 60, (b, o) => calc.createSession(b, o) as Promise<unknown> as Promise<Undoable>, { clock, execution })
  })

  it('seed 9, in a thread', async () => {
    const calc = await made({ threads: 1 })
    await undoRun(9, 60, (b, o) => calc.createSession(b, o) as Promise<unknown> as Promise<Undoable>, { clock, execution })
  }, 120_000)

  it('seed 10, in a thread lost now and then, its snapshots taken every 3 calls', async () => {
    const { calc, crash } = await inside({ threads: 1, snapshotEvery: 3 })
    keep(calc)
    await undoRun(10, 60, (b, o) => calc.createSession(b, o) as Promise<unknown> as Promise<Undoable>, { clock, execution }, (k) => k % 9 === 8 && crash(0))
  }, 120_000)
})
