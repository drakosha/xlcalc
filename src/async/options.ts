/** Checks of the asynchronous calls' and the calculators' options, as `input.ts` checks the rest (`InputError`). */
import type { AbortSignalLike, CalcLimits, CalcOptions, CalcPoolOptions, InlineCalcOptions } from '../contract.js'
import { checkLimits, InputError, isObject, onlyOptions } from '../input.js'

/** A call's limits over the calculator's, field by field: a field the call does not give keeps the calculator's (a pool's bounds by its heap too). */
export const callLimits = (calc: CalcLimits | undefined, call: CalcLimits | undefined): CalcLimits | undefined => (calc && call ? { ...calc, ...call } : (call ?? calc))

function options(v: unknown, path: string, allowed: readonly string[]): Record<string, unknown> {
  if (!isObject(v)) throw new InputError(path, `expected an options object, got ${v === null ? 'null' : typeof v}`)
  onlyOptions(v, path, allowed)
  return v
}

function number(v: unknown, path: string, min: number, integer: boolean) {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < min || (integer && !Number.isInteger(v))) throw new InputError(path, `expected ${integer ? 'an integer' : 'a number'} of at least ${min}, got ${JSON.stringify(v)}`)
}

function signal(v: unknown, path: string): AbortSignalLike {
  const s = v as AbortSignalLike
  if (!isObject(v) || typeof s.aborted !== 'boolean' || typeof s.addEventListener !== 'function' || typeof s.removeEventListener !== 'function') throw new InputError(path, 'expected an AbortSignal')
  return s
}

/** A call's options: its `signal` and `timeoutMs` checked, and the rest (compute or session options) as they are. */
export function splitCall<T extends object>(v: unknown, path = 'opts'): CalcOptions & { rest: T | undefined } {
  if (v === undefined) return { rest: undefined }
  if (!isObject(v)) throw new InputError(path, `expected an options object, got ${v === null ? 'null' : typeof v}`)
  const { signal: sig, timeoutMs, ...rest } = v
  if (timeoutMs !== undefined) number(timeoutMs, `${path}.timeoutMs`, 0, false)
  return { signal: sig === undefined ? undefined : signal(sig, `${path}.signal`), timeoutMs: timeoutMs as number | undefined, rest: rest as T }
}

/** A call made with its own options (`rest`) apart from its `signal` and `timeoutMs`; options it refuses as a rejected promise. */
export function withCall<T extends object, R>(v: unknown, go: (rest: T | undefined, call: CalcOptions) => Promise<R>): Promise<R> {
  try {
    const { signal, timeoutMs, rest } = splitCall<T>(v)
    return go(rest, { signal, timeoutMs })
  } catch (e) {
    return Promise.reject(e)
  }
}

/** The options of one session call: `signal` and `timeoutMs` only. */
export function checkCall(v: unknown, path = 'opts'): CalcOptions {
  if (v === undefined) return {}
  const o = options(v, path, ['signal', 'timeoutMs'])
  return splitCall(o, path)
}

/**
 * The options of a session call that changes the book: `signal` and `timeoutMs` for the call, checked here, and the
 * edit's (`baseRevision`, and `extra`), which the session checks when the call runs; undefined when there are none.
 */
export function checkEditCall(v: unknown, extra: readonly string[] = [], path = 'opts'): [CalcOptions, Record<string, unknown> | undefined] {
  if (v === undefined) return [{}, undefined]
  const o = options(v, path, ['signal', 'timeoutMs', 'baseRevision', ...extra])
  const { signal: _, timeoutMs: __, ...edit } = o
  const { rest: ___, ...call } = splitCall(o, path)
  return [call, Object.keys(edit).length ? edit : undefined]
}

export function checkInlineOptions(v: unknown, path = 'opts'): InlineCalcOptions {
  if (v === undefined) return {}
  const o = options(v, path, ['sliceMs', 'timeoutMs', 'limits'])
  if (o.sliceMs !== undefined) number(o.sliceMs, `${path}.sliceMs`, 0, false)
  if (o.timeoutMs !== undefined) number(o.timeoutMs, `${path}.timeoutMs`, 0, false)
  if (o.limits !== undefined) checkLimits(o.limits, `${path}.limits`)
  return o as InlineCalcOptions
}

const POOL_COUNTS = ['threads', 'maxHeapMb', 'maxQueue', 'maxSessions', 'maxSessionsPerThread', 'snapshotEvery'] as const
const POOL_TIMES = ['hardTimeoutMs', 'timeoutMs', 'sliceMs'] as const

export function checkPoolOptions(v: unknown, path = 'opts'): CalcPoolOptions {
  if (v === undefined) return {}
  const o = options(v, path, [...POOL_COUNTS, ...POOL_TIMES, 'limits', 'spawn'])
  for (const k of POOL_COUNTS) if (o[k] !== undefined) number(o[k], `${path}.${k}`, k === 'maxQueue' ? 0 : 1, true)
  for (const k of POOL_TIMES) if (o[k] !== undefined) number(o[k], `${path}.${k}`, 0, false)
  if (o.limits !== undefined) checkLimits(o.limits, `${path}.limits`)
  if (o.spawn !== undefined && typeof o.spawn !== 'function') throw new InputError(`${path}.spawn`, 'expected a function that starts a Worker')
  return o as CalcPoolOptions
}
