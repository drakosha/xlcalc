/**
 * The value model of the runtime: errors as values, scalars, range references, and the basic coercions
 * of operators and functions (number, text, boolean).
 */
import type { CellValue, ErrorCode } from './contract.js'
import { collator } from './criteria.js'
import { textToNumber } from './dates.js'
import { generalText } from './format.js'
import { NUMBER_TEXT, readNumber } from './model.js'

export class CErr {
  constructor(readonly code: ErrorCode) {}
}

export const E: Record<ErrorCode, CErr> = {
  DIV0: new CErr('DIV0'),
  NA: new CErr('NA'),
  NAME: new CErr('NAME'),
  NUM: new CErr('NUM'),
  REF: new CErr('REF'),
  VALUE: new CErr('VALUE'),
  CYCLE: new CErr('CYCLE'),
  SPILL: new CErr('SPILL'),
  CALC: new CErr('CALC'),
  ERROR: new CErr('ERROR'),
}

export type V = number | string | boolean | null | CErr

/** A sheet rectangle; a whole column is cut to the occupied rows of the sheet when read. */
export class Rng {
  /** Key for the value and criteria-index caches. */
  readonly key: string
  constructor(
    readonly s: number,
    readonly r1: number,
    readonly c1: number,
    readonly r2: number,
    readonly c2: number,
  ) {
    this.key = `${s}:${r1}:${c1}:${r2}:${c2}`
  }
}

export const isErr = (x: unknown): x is CErr => x instanceof CErr

/** The smallest positive normal double: Excel has no smaller numbers (subnormal results are 0). */
const MIN_NORMAL = 2.2250738585072014e-308

/** A computed number as Excel keeps it: a non-finite one is #NUM!, one below the smallest normal number is 0. */
export const fitNumber = (x: number): number | CErr => (!Number.isFinite(x) ? E.NUM : Math.abs(x) < MIN_NORMAL ? 0 : x)

export function toNum(v: V): number | CErr {
  if (typeof v === 'number') return v
  if (v === null) return 0
  if (typeof v === 'boolean') return v ? 1 : 0
  if (isErr(v)) return v
  // decimal text first (the common case), else as VALUE reads it: `$1,000`, `50%`, dates and times
  const t = v.trim()
  return NUMBER_TEXT.test(t) ? readNumber(t) : (textToNumber(v) ?? E.VALUE)
}

export function toText(v: V): string | CErr {
  if (typeof v === 'string') return v
  if (v === null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (isErr(v)) return v
  return generalText(v)
}

/** Text "TRUE" and "FALSE" are booleans, as in Excel en-US with the English UI. */
export function toBool(v: V): boolean | CErr {
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v === null) return false
  if (isErr(v)) return v
  const u = v.toUpperCase()
  if (u === 'TRUE') return true
  if (u === 'FALSE') return false
  return E.VALUE
}

const rank = (v: V) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)

/**
 * Excel's order of two values that are not errors: numbers < text < booleans, text case-insensitive
 * (`collator`); a blank is 0, "" or FALSE, as the other side requires.
 */
export function compareValues(a: V, b: V): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'string') return collator.compare(a, b as string)
  return (a as number) < (b as number) ? -1 : (a as number) > (b as number) ? 1 : 0
}

/** Whether two non-blank values are of the same type for comparison (number, text, boolean). */
export const sameType = (a: V, b: V) => rank(a) === rank(b)

export function fromCellValue(v: CellValue): V {
  return v !== null && typeof v === 'object' ? E[v.error] : v
}
