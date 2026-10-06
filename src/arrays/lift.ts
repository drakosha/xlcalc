/**
 * Lifting (`docs/arrays-incell.md`): an operator or a one-value argument applied to arrays is computed
 * element by element. The result has the larger rows and the larger columns of the operands; an operand
 * with one row (column) is repeated down (across), otherwise a position outside it is #N/A. Only the
 * occupied part of the operands is computed, the rest is one `fill` value.
 */
import { E, Rng, type V } from '../values.js'
import { Arr } from './value.js'

/** A value as the runtime passes it: a scalar, a reference or an array. */
export type Val = V | Rng | Arr

/** A reference or an array of more than one value (a 1 × 1 one is a scalar wherever it goes). */
export const isMany = (x: unknown): x is Rng | Arr =>
  x instanceof Arr ? x.rows * x.cols > 1 : x instanceof Rng && (x.r2 > x.r1 || x.c2 > x.c1)

/** Element (r, c) of an operand, a single row or column repeated; undefined past the operand. */
function elem(x: V | Arr, r: number, c: number): V | undefined {
  if (!(x instanceof Arr)) return x
  const rr = x.rows === 1 ? 0 : r
  const cc = x.cols === 1 ? 0 : c
  return rr >= x.rows || cc >= x.cols ? undefined : x.at(rr, cc)
}

/** How far an operand varies along one dimension of size `n`: its occupied part, all of it when repeated or short. */
function extent(size: number, used: number, filled: boolean, n: number): number {
  if (size === n) return used
  if (size === 1) return filled ? n : 0
  return n
}

/** `f` over the operands element by element (arrays among `xs` lifted, the rest repeated). */
export function liftN(xs: readonly (V | Arr)[], f: (args: V[]) => V): Arr {
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
  const out: V[] = new Array(h * w)
  const args: V[] = new Array(xs.length)
  for (let r = 0; r < h; r++)
    for (let c = 0; c < w; c++) {
      let outside = false
      for (let i = 0; i < xs.length; i++) {
        const v = elem(xs[i], r, c)
        if (v === undefined) outside = true
        else args[i] = v
      }
      out[r * w + c] = outside ? E.NA : f(args)
    }
  const fill = h < rows || w < cols ? f(xs.map((x) => (x instanceof Arr ? x.fill : x))) : null
  return new Arr(rows, cols, h, w, out, fill)
}

/**
 * The value of a lifted call at one position: an array or a reference answer gives its top-left value, and a
 * blank there stays an empty element (INDEX); a blank value becomes 0 (IF over a blank cell, as in Excel).
 */
export function first(x: Val, cell: (g: Rng) => V): V {
  if (x instanceof Arr) return x.at(0, 0)
  if (x instanceof Rng) return cell(x)
  return x === null ? 0 : x
}
