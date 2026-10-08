/** Checks of the asynchronous calls' and the calculators' options, as `input.ts` checks the rest (`InputError`). */
import type { AbortSignalLike, CalcOptions, CalcPoolOptions, InlineCalcOptions } from '../contract.js'
import { checkLimits, InputError } from '../input.js'

const isObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v)

function options(v: unknown, path: string, allowed: readonly string[]): Record<string, unknown> {
  if (!isObject(v)) throw new InputError(path, `expected an options object, got ${v === null ? 'null' : typeof v}`)
  for (const k of Object.keys(v)) if (!allowed.includes(k)) throw new InputError(path, `unknown field ${JSON.stringify(k)} (expected ${allowed.join(', ')})`)
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

/** The options of one session call: `signal` and `timeoutMs` only. */
export function checkCall(v: unknown, path = 'opts'): CalcOptions {
  if (v === undefined) return {}
  const o = options(v, path, ['signal', 'timeoutMs'])
  return splitCall(o, path)
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
