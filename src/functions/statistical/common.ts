/** Helpers shared by the statistical function files. */
import { pairs, type Grid } from '../../args.js'
import { E, isErr, type CErr } from '../../values.js'
import { momentValue, stableMean } from './moments.js'

export const mean = stableMean
export const sorted = (xs: readonly number[]) => [...xs].sort((a, b) => a - b)

/** Variance and standard deviation share the same degrees-of-freedom check across function families. */
const varianceValue = (xs: readonly number[], ddof: number, root: boolean) =>
  xs.length - ddof < 1 ? E.DIV0 : momentValue(xs, xs.length - ddof, root)
export const variance = (xs: readonly number[], ddof: number) => varianceValue(xs, ddof, false)
export const stdev = (xs: readonly number[], ddof: number) => varianceValue(xs, ddof, true)

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

/**
 * Rule 4 with errors, in Excel's order after a one-cell reference's own error (kind `C`): different sizes (#N/A), an
 * error in either grid, then pairs of numbers by position.
 */
export function numberPairs(a: Grid, b: Grid): [number[], number[]] | CErr {
  return a.rows * a.cols !== b.rows * b.cols ? E.NA : (gridError(a) ?? gridError(b) ?? pairs(a, b))
}
