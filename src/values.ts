/**
 * The value model of the runtime: errors as values, scalars, range references, and the basic coercions
 * of operators and functions (number, text, boolean).
 */
import type { CellValue, ErrorCode } from './contract.js'
import { collator, wordKey } from './collate.js'
import { decimalFirst, textToNumber } from './dates.js'
import { generalText } from './numfmt/general.js'

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
  NULL: new CErr('NULL'),
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

/**
 * A reference of several areas: a union on one sheet (`(A1:A3,C1:C3)`), or one rectangle on each sheet of a
 * reference across sheets (`'S1:S3'!A1:B2`, `cube`). Functions over lists of values take each area in turn; one value
 * cannot be read from it (#VALUE!).
 */
export class Areas {
  constructor(
    readonly list: readonly Rng[],
    readonly cube = false,
  ) {}
}

export const isErr = (x: unknown): x is CErr => x instanceof CErr

/** The smallest positive normal double: Excel has no smaller numbers (subnormal results are 0). */
export const MIN_NORMAL = 2.2250738585072014e-308

/** A computed number as Excel keeps it: a non-finite one is #NUM!, one below the smallest normal number is 0. */
export const fitNumber = (x: number): number | CErr => (!Number.isFinite(x) ? E.NUM : Math.abs(x) < MIN_NORMAL ? 0 : x)

/**
 * Excel's last addition of a formula or of SUM: its answer `r` is 0 below 2^(e-49), e the exponent of the left operand
 * `a` (2^e ≤ |a| < 2^(e+1)), eight units in the last place of a (the sets `near-zero-held`, `sum-zero-held`).
 */
export function nearZero(a: number, r: number): number {
  const x = Math.abs(a)
  const y = Math.abs(r)
  // an answer of 2^-48 of the operand or more is above the bound whatever the operand's mantissa
  if (y * 2 ** 48 >= x || y === 0) return r
  // the operand's power of two: log2 can round up to the next whole number just below a power of two
  let p = 2 ** Math.floor(Math.log2(x))
  if (p > x) p /= 2
  else if (p * 2 <= x) p *= 2
  return y * 2 ** 49 < p ? 0 : r
}

/**
 * The coercions of a locale: `toNum` (operators and numeric arguments; text through `number`), `toText`
 * (`&` and text arguments; booleans by `words`, numbers through `numberText`), `toBool` (text that is one of
 * `words`, any case).
 */
export function coercions(number: (text: string) => number | null, words: { readonly TRUE: string; readonly FALSE: string }, numberText: (n: number) => string) {
  const T = wordKey(words.TRUE)
  const F = wordKey(words.FALSE)
  return {
    toNum(v: V): number | CErr {
      if (typeof v === 'number') return v
      if (v === null) return 0
      if (typeof v === 'boolean') return v ? 1 : 0
      if (isErr(v)) return v
      return number(v) ?? E.VALUE
    },
    toText(v: V): string | CErr {
      if (typeof v === 'string') return v
      if (v === null) return ''
      if (typeof v === 'boolean') return v ? words.TRUE : words.FALSE
      if (isErr(v)) return v
      return numberText(v)
    },
    toBool(v: V): boolean | CErr {
      if (typeof v === 'boolean') return v
      if (typeof v === 'number') return v !== 0
      if (v === null) return false
      if (isErr(v)) return v
      const u = wordKey(v)
      return u === T ? true : u === F ? false : E.VALUE
    },
  }
}

/** en-US with the English UI outside a book: decimal text first (the common case), else as operators read it (`$1,000`, `50%`, dates and times); TRUE and FALSE. */
export const { toNum, toText, toBool } = coercions(decimalFirst(textToNumber), { TRUE: 'TRUE', FALSE: 'FALSE' }, generalText)

const rank = (v: V) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)

/**
 * Excel's order of two values that are not errors, with `text` the order of text (case-insensitive, by
 * locale): numbers < text < booleans; a blank is 0, "" or FALSE, as the other side requires.
 */
export function compareWith(text: (a: string, b: string) => number): (a: V, b: V) => number {
  return (a, b) => {
    if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
    if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0
    const ra = rank(a)
    const rb = rank(b)
    if (ra !== rb) return ra - rb
    if (typeof a === 'string') return text(a, b as string)
    return (a as number) < (b as number) ? -1 : (a as number) > (b as number) ? 1 : 0
  }
}

/** Excel en-US's order of two values that are not errors (`collator`). */
export const compareValues = compareWith(collator.compare)

/** A number to 15 significant digits, as Excel compares numbers and reads them in INT and its kin (`1-2^-52` is 1 there). */
export const to15 = (x: number): number => (Number.isInteger(x) && Math.abs(x) < 1e15 ? x : Number(x.toPrecision(15)))

/**
 * The order of the comparison operators (`=`, `<` and the rest) over `compare`: numbers equal to 15 significant
 * digits are equal (`0.1+0.2=0.3` is TRUE), as in Excel; lookups, sorting and UNIQUE compare exactly.
 */
export const operatorOrder =
  (compare: (a: V, b: V) => number) =>
  (a: V, b: V): number =>
    compare(typeof a === 'number' ? to15(a) : a, typeof b === 'number' ? to15(b) : b)

/** The comparison operators' order in en-US. */
export const compareOp = operatorOrder(compareValues)

/** Whether a comparison operator (`=`, `<>`, `<`, `>`, `<=`, `>=`) holds for `c`, the order of its operands. */
export const holds = (op: string, c: number) => (op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0)

/** Excel's tolerance of a whole-number argument: a number less than 2^-22 short of an integer is that integer. */
export const WHOLE = 2 ** -22

/**
 * A whole-number argument as Excel reads most of them (LEFT, INDEX, ROUND's digits, DATE, measured by the
 * `near-integer` set): toward zero after adding the tolerance to the magnitude (`LEFT("abc",0.9999999)` is "a").
 */
export const whole = (x: number, tolerance = WHOLE): number => (x < 0 ? 0 - Math.trunc(tolerance - x) : Math.trunc(x + tolerance))

/** Whether two non-blank values are of the same type for comparison (number, text, boolean). */
export const sameType = (a: V, b: V) => rank(a) === rank(b)

export function fromCellValue(v: CellValue): V {
  return v !== null && typeof v === 'object' ? E[v.error] : v
}
