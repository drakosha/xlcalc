/**
 * Excel en-US cases for the largest number of arguments of every whitelisted function with repeated
 * arguments: a call with that many arguments, which Excel loads. Excel refuses a book with one more (the
 * "We found a problem with some content" window, closed by the harness), so those are not cases; the limits
 * are in `LIMITS`, which `test/arity.test.ts` checks against the engine. Computed as dynamic array formulas, in
 * Excel en-US (`test/excel/cases/arity.json`).
 */
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'

interface Shape {
  /** The arguments before the repeated ones. */
  head: string[]
  /** One repetition of the repeated arguments. */
  rep: string[]
  /** The arguments after them (SWITCH's default). */
  tail?: string[]
  /** A function with an array answer is wrapped (ROWS) to give one value. */
  wrap?: string
}

const ones = (rep = ['1']): Shape => ({ head: [], rep })
const crit = (head: string[]): Shape => ({ head, rep: ['D!C1:C8', '">0"'] })

/** How a call of each function is built. */
export const SHAPES: Record<string, Shape> = {
  ...Object.fromEntries(
    ['IMPRODUCT', 'IMSUM', 'GCD', 'LCM', 'MULTINOMIAL', 'PRODUCT', 'SUM', 'SUMSQ', 'AVEDEV', 'AVERAGE', 'AVERAGEA', 'COUNT', 'COUNTA', 'DEVSQ', 'GEOMEAN', 'HARMEAN', 'KURT', 'MAX', 'MAXA', 'MEDIAN', 'MIN', 'MINA', 'MODE', 'MODE.SNGL', 'SKEW', 'SKEW.P', 'STDEV', 'STDEV.P', 'STDEV.S', 'STDEVA', 'STDEVP', 'STDEVPA', 'VAR', 'VAR.P', 'VAR.S', 'VARA', 'VARP', 'VARPA'].map((f) => [f, ones()]),
  ),
  'MODE.MULT': { ...ones(), wrap: 'ROWS' },
  NPV: { head: ['0.1'], rep: ['1'] },
  AND: ones(['TRUE']),
  OR: ones(['TRUE']),
  XOR: ones(['TRUE']),
  IFS: { head: [], rep: ['TRUE', '1'] },
  SWITCH: { head: ['1'], rep: ['2'], tail: ['1'] },
  CHOOSE: { head: ['1'], rep: ['1'] },
  CHOOSECOLS: { head: ['D!B1:C8'], rep: ['1'], wrap: 'ROWS' },
  CHOOSEROWS: { head: ['D!B1:C8'], rep: ['1'], wrap: 'ROWS' },
  HSTACK: { ...ones(), wrap: 'ROWS' },
  VSTACK: { ...ones(), wrap: 'ROWS' },
  SORTBY: { head: ['D!B1:B8'], rep: ['D!C1:C8', '1'], wrap: 'ROWS' },
  AGGREGATE: { head: ['9', '0'], rep: ['D!B1:B8'] },
  SUBTOTAL: { head: ['9'], rep: ['D!B1:B8'] },
  SUMPRODUCT: ones(['D!B1:B8']),
  SUMIFS: crit(['D!B1:B8']),
  AVERAGEIFS: crit(['D!B1:B8']),
  MAXIFS: crit(['D!B1:B8']),
  MINIFS: crit(['D!B1:B8']),
  COUNTIFS: crit([]),
  CONCAT: ones(['"a"']),
  CONCATENATE: ones(['"a"']),
  TEXTJOIN: { head: ['","', 'TRUE'], rep: ['"a"'] },
}

/** The call of `fn` with `n` arguments; null when no whole number of repetitions gives `n`. */
export function call(fn: string, n: number): string | null {
  const s = SHAPES[fn]
  const fixed = s.head.length + (s.tail?.length ?? 0)
  const k = (n - fixed) / s.rep.length
  if (!Number.isInteger(k) || k < 1) return null
  const args = [...s.head, ...Array.from({ length: k }, () => s.rep).flat(), ...(s.tail ?? [])]
  const f = `${fn}(${args.join(',')})`
  return s.wrap ? `${s.wrap}(${f})` : f
}

/**
 * Excel's largest number of arguments where it is not 255: the functions newer than Excel 2007 (stored with the
 * `_xlfn.` prefix) take 254, or 253 where the repeated arguments come in pairs, and COUNTIFS's pairs end at 254.
 * Excel refused a book with a call of 255 arguments of each of the newer functions and loaded every other
 * function with 255 (measured on 7 October 2026, two runs).
 */
export const LIMITS: Record<string, number> = {
  ...Object.fromEntries(['CHOOSECOLS', 'CHOOSEROWS', 'CONCAT', 'HSTACK', 'IFS', 'MODE.MULT', 'MODE.SNGL', 'SKEW.P', 'STDEV.P', 'STDEV.S', 'SWITCH', 'TEXTJOIN', 'VAR.P', 'VAR.S', 'VSTACK', 'XOR', 'AGGREGATE'].map((f) => [f, 254])),
  COUNTIFS: 254,
  MAXIFS: 253,
  MINIFS: 253,
  SORTBY: 253,
}

export const limit = (fn: string) => LIMITS[fn] ?? 255

export function buildProbes(): Probe[] {
  return Object.keys(SHAPES).map((fn) => ({ id: `${fn}:${limit(fn)}`, group: fn, fn, formula: xl(call(fn, limit(fn))!) }))
}
