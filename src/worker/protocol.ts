/**
 * The messages between the pool and its threads (`docs/async.md`). Everything goes by structured clone; an error
 * travels as plain data and is made a `CalcError` or an `InputError` again on the other side.
 */
import type { BookInput, CalcLimits, CellRef, Clock, SessionHistory, SessionOptions } from '../contract.js'
import { CalcError } from '../exec/budget.js'
import { InputError } from '../input.js'
import type { HistoryData } from '../session/undo.js'
import type { Sent } from './pack.js'

export type { Sent }

/** The session calls that change the session: they go into its journal. */
export const CHANGING = ['setCells', 'replaceCells', 'clearCells', 'patchCells', 'setCellText', 'setNames', 'setClock', 'restructure', 'setStyles', 'setLayout', 'undo', 'redo'] as const
export type Op = (typeof CHANGING)[number] | 'previewCellText' | 'book' | 'layout' | 'extent' | 'explain' | 'cells' | 'getViewport' | 'getCellIssue' | 'validation' | 'validateEntry' | 'result' | 'getSnapshot' | 'hash'

/** The calls that take edit options (`EditOptions`), by the number of arguments before them: their messages carry the options last. */
export const EDITS: Partial<Record<Op, number>> = { setCells: 1, replaceCells: 1, clearCells: 2, patchCells: 1, setCellText: 3, previewCellText: 3, setNames: 1, setClock: 1, restructure: 1, setStyles: 1, setLayout: 1, undo: 0, redo: 0 }

/** A call the session answered, as its journal keeps it: replayed with the clock the session had after it. */
export interface Entry {
  seq: number
  op: Op
  args: unknown[]
  clock: Clock
}


/** A session's book at sequence number `seq`: what a restore opens before it replays the journal. */
export interface Snapshot {
  book: Sent
  /** The session's options without the clock. */
  opts: SessionOptions
  clock: Clock
  seq: number
  /** The session's history then (as text when JSON holds it): the replayed `undo` and `redo` calls take its steps. */
  history?: HistoryData | { json: string }
  /** The session's revision then. */
  revision?: number
}

/** The book's outline before a structure operation: enough for `applyOp` to give the same `mapCell`. */
export type Outline = BookInput

/** To a thread. A thread runs one task at a time; `cancel` may come while it runs. */
export type ToThread =
  /** `cacheBytes`: what the range caches of the thread's books may hold together (`docs/cache.md`). */
  | { t: 'init'; sliceMs: number; limits?: CalcLimits; cacheBytes?: number }
  | { t: 'compute'; task: number; book: Sent; opts: object; timeoutMs?: number }
  | { t: 'open'; task: number; sid: number; book: Sent; opts: SessionOptions; timeoutMs?: number }
  | { t: 'call'; task: number; sid: number; seq?: number; op: Op; args: unknown[]; timeoutMs?: number }
  | { t: 'restore'; task: number; sid: number; snapshot: Snapshot; journal: Entry[] }
  | { t: 'snapshot'; task: number; sid: number }
  /** `keep`: the session leaves a full thread, its book comes back as its new snapshot. */
  | { t: 'drop'; task: number; sid: number; keep?: boolean }
  | { t: 'cancel'; task: number }

export interface ErrorData {
  name: string
  message: string
  code?: string
  path?: string
  limit?: keyof CalcLimits
  max?: number
  size?: number
  cell?: CellRef
  cause?: unknown
}

/** From a thread: a task's answer. */
export type FromThread = { task: number; ok: true; value: unknown } | { task: number; ok: false; error: ErrorData }

/** A call's value and the session's clock, history and revision after it. */
export interface Answer {
  value: unknown
  clock: Clock
  history: SessionHistory
  revision: number
  /** `restructure`: the book's outline before it. */
  outline?: Outline
}

const plain = (v: unknown): unknown => {
  if (v === undefined || v === null || ['string', 'number', 'boolean'].includes(typeof v)) return v
  return String(v)
}

export function toData(e: unknown): ErrorData {
  if (e instanceof CalcError) return { name: e.name, message: e.message, code: e.code, limit: e.limit, max: e.max, size: e.size, cell: e.cell, cause: plain(e.cause) }
  if (e instanceof InputError) return { name: e.name, message: e.message, path: e.path }
  if (e instanceof Error) return { name: e.name, message: e.message }
  return { name: 'Error', message: String(e) }
}

export function fromData(d: ErrorData): Error {
  if (d.name === 'CalcError') {
    const { limit, max, size, cell, cause } = d
    return new CalcError(d.code as CalcError['code'], d.message, { limit, max, size, cell, cause })
  }
  if (d.name === 'InputError' && d.path !== undefined) return new InputError(d.path, d.message.slice(d.path.length + 2))
  const e = new Error(d.message)
  e.name = d.name
  return e
}
