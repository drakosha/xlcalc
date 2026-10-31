/**
 * The 11 functions of a value field over the records at a cell of the cube. A cell no value reaches (every record blank
 * in the field) is empty; an error among the values is the answer but for `count` (which counts it) and `countNums`;
 * text and booleans count for `count` alone, so a sum over them is 0, an average #DIV/0!. A total is made of the tallies
 * of the cells under it or of their records, in the order supplied by the calculation. Numeric operations follow
 * Excel's two rounding steps; error precedence follows the cache groups and their records.
 */
import { pivotAdd, pivotMultiply, pivotDivide, pivotRoot } from './number.js'
import type { CellValue, ErrorCode, PivotFunction } from '../contract.js'

/** The running state of one value field at one cell. */
export class Tally {
  /** Values not blank, errors among them. */
  filled = 0
  numbers = 0
  sum = 0
  /** The sum of the numbers' squares. */
  squares = 0
  product = 1
  min = Infinity
  max = -Infinity
  error?: ErrorCode
  private errorAt?: readonly number[]

  /** `lastError`: the last error is the answer, as a calculated field takes a field's sum; the first without it. */
  constructor(private readonly lastError = false) {}

  /** Error precedence follows cache groups, then source records, even when numbers traverse displayed groups. */
  private takeError(error: ErrorCode | undefined, at?: readonly number[]): void {
    if (!error) return
    const d = at && this.errorAt ? at.findIndex((n, i) => n !== this.errorAt![i]) : -1
    const replace = this.error === undefined || (at && this.errorAt
      ? d < 0 ? this.lastError : this.lastError ? at[d] > this.errorAt[d] : at[d] < this.errorAt[d]
      : this.lastError)
    if (replace) { this.error = error; this.errorAt = at }
  }

  /** A record's value: a number (a date's serial), other non-blank, an error, or blank (`filled` false). */
  add(n: number | undefined, filled: boolean, error?: ErrorCode, at?: readonly number[]): void {
    if (!filled) return
    this.filled++
    if (error) {
      this.takeError(error, at)
      return
    }
    if (n === undefined) return
    this.numbers++
    this.sum = pivotAdd(this.sum, n)
    this.squares = pivotAdd(this.squares, pivotMultiply(n, n))
    this.product = pivotMultiply(this.product, n)
    if (n < this.min) this.min = n
    if (n > this.max) this.max = n
  }

  /** Takes in the tally of a cell under this one. */
  merge(t: Tally): void {
    this.filled += t.filled
    this.takeError(t.error, t.errorAt)
    if (!t.numbers) return
    this.numbers += t.numbers
    this.sum = pivotAdd(this.sum, t.sum)
    this.squares = pivotAdd(this.squares, t.squares)
    this.product = pivotMultiply(this.product, t.product)
    if (t.min < this.min) this.min = t.min
    if (t.max > this.max) this.max = t.max
  }

  /** The cell's value by `fn`. */
  value(fn: PivotFunction): CellValue {
    if (!this.filled) return null
    const n = this.numbers
    if (fn === 'count') return this.filled
    if (fn === 'countNums') return n
    if (this.error) return { error: this.error }
    const { sum, squares } = this
    switch (fn) {
      case 'sum':
        return sum
      case 'average':
        return n ? pivotDivide(sum, n) : { error: 'DIV0' }
      case 'max':
        return n ? this.max : 0
      case 'min':
        return n ? this.min : 0
      case 'product':
        return n ? this.product : 0
    }
    const sample = fn === 'stdDev' || fn === 'var'
    if (n - (sample ? 1 : 0) < 1) return { error: 'DIV0' }
    // Excel's own forms: (Σx² − (Σx)²/n)/(n − 1) for a sample, (nΣx² − (Σx)²)/n² for a population
    const square = pivotMultiply(sum, sum)
    const v = sample ? pivotDivide(pivotAdd(squares, -pivotDivide(square, n)), n - 1) : pivotDivide(pivotAdd(pivotMultiply(n, squares), -square), pivotMultiply(n, n))
    const out = fn === 'stdDev' || fn === 'stdDevp' ? pivotRoot(Math.max(v, 0)) : Math.max(v, 0)
    return Number.isFinite(out) ? out : { error: 'NUM' }
  }
}
