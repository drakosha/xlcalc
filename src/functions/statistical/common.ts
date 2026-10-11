/** Helpers shared by the statistical function files. */
import { pairs, type Arg, type Cx, type Grid } from '../../args.js'
import { Areas, E, isErr, type CErr, type V } from '../../values.js'
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

/**
 * The numbers among values in order, or among those at the positions `at` (a criteria selection): text, booleans and
 * blanks are skipped, the first error is the answer.
 */
export function numbersOf(values: readonly V[], at?: readonly number[]): number[] | CErr {
  const out: number[] = []
  const n = at ? at.length : values.length
  for (let i = 0; i < n; i++) {
    const v = values[at ? at[i] : i]
    if (typeof v === 'number') out.push(v)
    else if (isErr(v)) return v
  }
  return out
}

/** The numbers of an argument's grid, of each area in turn for several (a union, a reference across sheets); an error is the answer. */
export function areaNumbers(cx: Cx, x: Arg): number[] | CErr {
  if (isErr(x)) return x
  const out: number[] = []
  for (const g of x instanceof Areas ? x.list : [x]) {
    const xs = numbersOf(cx.grid(g).v)
    if (isErr(xs)) return xs
    // one push per number: a whole column would overflow the call stack
    for (const n of xs) out.push(n)
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
