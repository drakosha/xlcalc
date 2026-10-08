/**
 * The messages between the pool and its threads (`docs/async.md`). Everything goes by structured clone; an error
 * travels as plain data and is made a `CalcError` or an `InputError` again on the other side.
 */
import type { BookInput, CalcLimits, CellRef, Clock, SessionOptions } from '../contract.js'
import { CalcError } from '../exec/budget.js'
import { InputError } from '../input.js'

/** The session calls that change the session: they go into its journal. */
export const CHANGING = ['setCells', 'setNames', 'setClock', 'restructure', 'setStyles', 'setLayout'] as const
export type Op = (typeof CHANGING)[number] | 'book' | 'explain' | 'result'

/** A call the session answered, as its journal keeps it: replayed with the clock the session had after it. */
export interface Entry {
  seq: number
  op: Op
  args: unknown[]
  clock: Clock
}

/** A session's book at sequence number `seq`: what a restore opens before it replays the journal. */
export interface Snapshot {
  book: BookInput
  /** The session's options without the clock. */
  opts: SessionOptions
  clock: Clock
  seq: number
}

/** The book's outline before a structure operation: enough for `applyOp` to give the same `mapCell`. */
export type Outline = BookInput

/** To a thread. A thread runs one task at a time; `cancel` may come while it runs. */
export type ToThread =
  | { t: 'init'; sliceMs: number; limits?: CalcLimits }
  | { t: 'compute'; task: number; book: BookInput; opts: object; timeoutMs?: number }
  | { t: 'open'; task: number; sid: number; book: BookInput; opts: SessionOptions; timeoutMs?: number }
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
  limit?: 'formulas' | 'cells'
  max?: number
  size?: number
  cell?: CellRef
  cause?: unknown
}

/** From a thread: a task's answer. */
export type FromThread = { task: number; ok: true; value: unknown } | { task: number; ok: false; error: ErrorData }

/** A call's value and the session's clock after it. */
export interface Answer {
  value: unknown
  clock: Clock
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
