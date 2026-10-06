/** Helpers shared by the statistical function files. */
import { pairs, type Grid } from '../../args.js'
import { isErr, type CErr } from '../../values.js'

export const sum = (xs: readonly number[]) => xs.reduce((s, x) => s + x, 0)
export const mean = (xs: readonly number[]) => sum(xs) / xs.length
// loops, not Math.max(...xs): a whole column would overflow the call stack
export const max = (xs: readonly number[]) => xs.reduce((m, x) => (x > m ? x : m), -Infinity)
export const min = (xs: readonly number[]) => xs.reduce((m, x) => (x < m ? x : m), Infinity)
export const sorted = (xs: readonly number[]) => [...xs].sort((a, b) => a - b)

/** Σ(x − mean)², two passes as Excel. */
export function devsq(xs: readonly number[]): number {
  const m = mean(xs)
  let s = 0
  for (const x of xs) s += (x - m) * (x - m)
  return s
}

/** Σ(x − x̄)(y − ȳ). */
export function codev(xs: readonly number[], ys: readonly number[]): number {
  const mx = mean(xs)
  const my = mean(ys)
  let s = 0
  for (let i = 0; i < xs.length; i++) s += (xs[i] - mx) * (ys[i] - my)
  return s
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

/** Rule 4 with errors: an error in either grid is the answer, then pairs of numbers by position. */
export function numberPairs(a: Grid, b: Grid): [number[], number[]] | CErr {
  return gridError(a) ?? gridError(b) ?? pairs(a, b)
}
