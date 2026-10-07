/**
 * The budget of one computation (`docs/async.md`): a soft deadline, an abort signal and size limits, checked at points
 * inside the engine (`spend`, `allot`, `count`), and the slice of computing after which a step yields (`due`). The
 * computation being run sets its budget active; on the synchronous path without options none is, and each check is
 * one comparison. The engine's loops are generators (`Steps`): `run` takes one to its end at once.
 */
import type { AbortSignalLike, CalcErrorCode, CalcLimits, CellRef } from '../contract.js'
import { nowMs } from '../env.js'

/** Why an asynchronous or limited call failed, by `code`; a `limit` names the limit, its maximum, the size and the formula's cell. */
export class CalcError extends Error {
  readonly code: CalcErrorCode
  // declared only: a field without a value is no property of the error
  declare limit?: 'formulas' | 'cells'
  declare max?: number
  declare size?: number
  declare cell?: CellRef
  constructor(code: CalcErrorCode, message: string, details: { limit?: 'formulas' | 'cells'; max?: number; size?: number; cell?: CellRef; cause?: unknown } = {}) {
    const { cause, ...rest } = details
    super(message, cause === undefined ? undefined : { cause })
    this.name = 'CalcError'
    this.code = code
    for (const [k, v] of Object.entries(rest)) if (v !== undefined) Object.assign(this, { [k]: v })
  }
}

/** A computation in steps: it yields between steps (formulas) when its slice is used up. */
export type Steps<T> = Generator<void, T, void>

/** Work units between two looks at the clock and the signal. */
const QUANTUM = 1024

export class Budget {
  /** Work units left before the next look. */
  left = QUANTUM
  /** The slice is used up: the next step boundary yields. */
  due = false
  /** When the current slice ends (`nowMs`). */
  sliceEnd = Infinity
  readonly cells: number
  readonly formulas: number

  constructor(
    limits: CalcLimits = {},
    /** The soft deadline (`nowMs`), set when the computation starts. */
    public deadline = Infinity,
    readonly signal?: AbortSignalLike,
    readonly sliceMs = Infinity,
  ) {
    this.cells = limits.cells ?? Infinity
    this.formulas = limits.formulas ?? Infinity
  }

  /** Set by the host to stop the computation (its calculator closed). */
  halted: CalcError | null = null

  /** The reason to stop now, or null. */
  stopped(now = nowMs()): CalcError | null {
    if (this.halted) return this.halted
    if (this.signal?.aborted) return new CalcError('aborted', 'the computation was aborted', { cause: this.signal.reason })
    if (now > this.deadline) return new CalcError('soft_timeout', 'the computation ran past its timeoutMs')
    return null
  }

  look() {
    this.left = QUANTUM
    // a budget of limits alone needs no clock
    const now = this.deadline === Infinity && this.sliceEnd === Infinity ? 0 : nowMs()
    const stop = this.stopped(now)
    if (stop) throw stop
    if (now >= this.sliceEnd) this.due = true
  }

  /** A new slice of computing begins. */
  slice() {
    this.due = false
    this.left = QUANTUM
    this.sliceEnd = nowMs() + this.sliceMs
  }
}

let active: Budget | null = null

/** Counts `n` units of work: now and then stops the computation at its deadline or abort, or marks its slice used up. */
export function spend(n: number): void {
  if (active !== null && (active.left -= n) <= 0) active.look()
}

/** The computation's slice is used up: a step boundary yields. */
export const due = (): boolean => active !== null && active.due

/** `n` values are about to be made: over `limits.cells` stops the computation, before they are allocated. */
export function allot(n: number): void {
  if (active !== null && n > active.cells) throw new CalcError('limit', `${n} values in one range or array, over the limit of ${active.cells}`, { limit: 'cells', max: active.cells, size: n })
}

/** The book holds `n` formulas: over `limits.formulas` stops the computation. */
export function count(n: number): void {
  if (active !== null && n > active.formulas) throw new CalcError('limit', `${n} formulas, over the limit of ${active.formulas}`, { limit: 'formulas', max: active.formulas, size: n })
}

/** A limit error without a cell gets the cell of the formula being computed. */
export function located(e: unknown, cell: () => CellRef): unknown {
  if (e instanceof CalcError && e.code === 'limit' && e.limit === 'cells' && !e.cell) {
    e.cell = cell()
    e.message += ` (${e.cell.sheet}!${e.cell.cell})`
  }
  return e
}

/** `f` with `b` as the active budget (none: the one active stays). */
export function withBudget<T>(b: Budget | null | undefined, f: () => T): T {
  if (!b) return f()
  const outer = active
  active = b
  try {
    return f()
  } finally {
    active = outer
  }
}

/** Runs steps to their end at once. */
export function run<T>(g: Steps<T>): T {
  for (;;) {
    const r = g.next()
    if (r.done) return r.value
  }
}

/** One portion of steps with `b` active: the result when they end, or `undefined` done when they yielded. */
export function portion<T>(g: Steps<T>, b: Budget, how: 'next' | 'throw' = 'next', error?: unknown): IteratorResult<void, T> {
  const outer = active
  active = b
  b.slice()
  try {
    return how === 'next' ? g.next() : g.throw(error)
  } finally {
    active = outer
  }
}

/** A budget of limits alone (the synchronous path), or none without them. */
export const limitsBudget = (limits: CalcLimits | undefined): Budget | null => (limits && (limits.cells !== undefined || limits.formulas !== undefined) ? new Budget(limits) : null)
