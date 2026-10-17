/**
 * Lifting (`docs/arrays-incell.md`): an operator or a one-value argument applied to arrays is computed
 * element by element. The result has the larger rows and the larger columns of the operands; an operand
 * with one row (column) is repeated down (across), otherwise a position outside it is #N/A (`liftN`). Only the
 * occupied part of the operands is computed, the rest is one `fill` value.
 */
import { allot, spend } from '../exec/budget.js'
import { Areas, E, Rng, type V } from '../values.js'
import { Arr, broadcastAt } from './value.js'

/** A value as the runtime passes it: a scalar, a reference or an array. */
export type Val = V | Rng | Arr

/** A reference of more than one cell. */
export const isArea = (x: unknown): x is Rng => x instanceof Rng && (x.r2 > x.r1 || x.c2 > x.c1)

/** A reference or an array of more than one value: what one cell cannot show. */
export const isMany = (x: unknown): x is Rng | Arr => (x instanceof Arr ? x.rows * x.cols > 1 : isArea(x))

/** What an operator or a one-value argument lifts over: any array (a 1 × 1 one stays an array, `TYPE({1}+0)` is 64) or a reference of several cells. */
export const lifts = (x: unknown): x is Rng | Arr => x instanceof Arr || isArea(x)

/** How far an operand varies along one dimension of size `n`: its occupied part, all of it when repeated or short. */
function extent(size: number, used: number, filled: boolean, n: number): number {
  if (size === n) return used
  if (size === 1) return filled ? n : 0
  return n
}

/**
 * `f` over the operands element by element (arrays among `xs` lifted, the rest repeated). A position outside a
 * shorter array is #N/A for an operator; a function (`missing`) gets #N/A there as that argument's value and reads it
 * in its turn (`IFERROR(G4:G7/H4:H7,G12:G13)` keeps G6/H6, `LEFT(A1:A9,SEQUENCE(5))` gives A6's error).
 */
export function liftN(xs: readonly (V | Arr)[], f: (args: V[]) => V, missing = false): Arr {
  let rows = 1
  let cols = 1
  for (const x of xs)
    if (x instanceof Arr) {
      if (x.rows > rows) rows = x.rows
      if (x.cols > cols) cols = x.cols
    }
  let h = 0
  let w = 0
  for (const x of xs)
    if (x instanceof Arr) {
      const filled = x.h * x.w > 0
      h = Math.max(h, extent(x.rows, x.h, filled, rows))
      w = Math.max(w, extent(x.cols, x.w, filled, cols))
    }
  allot(h * w)
  const out: V[] = new Array(h * w)
  const args: V[] = new Array(xs.length)
  for (let r = 0; r < h; r++) {
    spend(w)
    for (let c = 0; c < w; c++) {
      let outside = false
      for (let i = 0; i < xs.length; i++) {
        const v = broadcastAt(xs[i], r, c)
        if (v !== undefined) args[i] = v
        else if (missing) args[i] = E.NA
        else outside = true
      }
      out[r * w + c] = outside ? E.NA : f(args)
    }
  }
  const fill = h < rows || w < cols ? f(xs.map((x) => (x instanceof Arr ? x.fill : x))) : null
  return new Arr(rows, cols, h, w, out, fill)
}

/**
 * The value of a lifted call at one position: an array or a reference answer gives its top-left value, and a
 * blank there stays an empty element (INDEX); a blank value becomes 0 (IF over a blank cell, as in Excel).
 */
export function first(x: Val | Areas, cell: (g: Rng) => V): V {
  if (x instanceof Arr) return x.at(0, 0)
  if (x instanceof Areas) return E.VALUE
  if (x instanceof Rng) return cell(x)
  return x === null ? 0 : x
}
