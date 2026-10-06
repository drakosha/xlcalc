/**
 * The array value of the runtime (`docs/arrays-incell.md`): rows × cols, an occupied part h × w from the
 * top-left corner stored row by row, and one `fill` value everywhere else, so that a whole column costs
 * only its occupied rows. Elements are any values, blanks (null) and errors included.
 */
import { E, isErr, type CErr, type V } from '../values.js'

export class Arr {
  constructor(
    readonly rows: number,
    readonly cols: number,
    readonly h: number,
    readonly w: number,
    /** h × w values row by row; shared with caches, never to be modified. */
    readonly v: readonly V[],
    readonly fill: V = null,
  ) {}

  at(r: number, c: number): V {
    return r < this.h && c < this.w ? this.v[r * this.w + c] : this.fill
  }

  /** Every position stored (h = rows, w = cols): for functions that read `v` as the whole array. */
  full(): Arr {
    if (this.h === this.rows && this.w === this.cols) return this
    const out: V[] = new Array(this.rows * this.cols)
    for (let r = 0; r < this.rows; r++) for (let c = 0; c < this.cols; c++) out[r * this.cols + c] = this.at(r, c)
    return new Arr(this.rows, this.cols, this.rows, this.cols, out)
  }

  /** The values row by row, every position (for list arguments). */
  values(): readonly V[] {
    return this.full().v
  }
}

export const isArr = (x: unknown): x is Arr => x instanceof Arr

/** A fully stored array from its values row by row. */
export const arr = (rows: number, cols: number, v: V[]): Arr => new Arr(rows, cols, rows, cols, v)

/** An array from rows of values (an array constant, a function's answer). */
export function fromRows(rows: V[][]): Arr {
  const cols = rows[0].length
  return arr(rows.length, cols, rows.flat())
}

/** A grid (a range's values, a scalar) as an array; an array stays as it is. */
export const fromGrid = (g: { rows: number; cols: number; h: number; w: number; v: readonly V[] }): Arr => (g instanceof Arr ? g : new Arr(g.rows, g.cols, g.h, g.w, g.v))

/** Excel has no empty arrays: an empty answer is #CALC!, which the contract has no code for (#VALUE!). */
export const EMPTY: CErr = E.VALUE

/** A 1 × 1 array is its single value; any other array stays. */
export const collapse = (a: Arr): V | Arr => (a.rows === 1 && a.cols === 1 ? a.at(0, 0) : a)

/** The first error of an array, row by row, or undefined. */
export function firstError(a: Arr): CErr | undefined {
  for (const v of a.v) if (isErr(v)) return v
  if (a.h * a.w < a.rows * a.cols && isErr(a.fill)) return a.fill
}
