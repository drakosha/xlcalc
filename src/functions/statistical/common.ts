/** Helpers shared by the statistical function files. */
import { pairs, type Grid } from '../../args.js'
import { E, isErr, type CErr } from '../../values.js'
import { momentValue, stableMean } from './moments.js'

export const sum = (xs: readonly number[]) => xs.reduce((s, x) => s + x, 0)
export const mean = stableMean
// loops, not Math.max(...xs): a whole column would overflow the call stack
export const max = (xs: readonly number[]) => xs.reduce((m, x) => (x > m ? x : m), -Infinity)
export const min = (xs: readonly number[]) => xs.reduce((m, x) => (x < m ? x : m), Infinity)
export const sorted = (xs: readonly number[]) => [...xs].sort((a, b) => a - b)

/** Σ(x − mean)², with corrected moments and an exact fallback. */
export function devsq(xs: readonly number[]): number {
  return momentValue(xs, 1)
}

/** The first error in the referenced part of a grid. */
export const gridError = (g: Grid): CErr | undefined => g.v.find(isErr) as CErr | undefined

/** The numbers of a grid; an error inside is the answer. */
export function gridNumbers(g: Grid): number[] | CErr {
  const out: number[] = []
  for (const v of g.v) {
    if (isErr(v)) return v
    if (typeof v === 'number') out.push(v)
  }
  return out
}

/** A grid of one cell holding an error: Excel reads it as the argument's own error, before any check of the function. */
export const cellError = (g: Grid): CErr | undefined => (g.rows * g.cols === 1 ? gridError(g) : undefined)

/** Rule 4 with errors, in Excel's order: a one-cell error, different sizes (#N/A), an error in either grid, then pairs of numbers by position. */
export function numberPairs(a: Grid, b: Grid): [number[], number[]] | CErr {
  return cellError(a) ?? cellError(b) ?? (a.rows * a.cols !== b.rows * b.cols ? E.NA : (gridError(a) ?? gridError(b) ?? pairs(a, b)))
}
