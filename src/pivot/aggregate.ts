/**
 * The 11 functions of a value field over the records at a cell of the cube. A cell no value reaches (every record blank
 * in the field) is empty; an error among the values is the answer but for `count` (which counts it) and `countNums`;
 * text and booleans count for `count` alone, so a sum over them is 0, an average #DIV/0!. A total is made of the tallies
 * of the cells under it, as Excel adds them (`merge`): the sums of a grand total are those of its lines added in turn,
 * not of the records, and the variances are Excel's sums of squares (the set `basic`).
 */
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

  /** `lastError`: the last error is the answer, as a calculated field takes a field's sum; the first without it. */
  constructor(private readonly lastError = false) {}

  /** A record's value: a number (a date's serial), other non-blank, an error, or blank (`filled` false). */
  add(n: number | undefined, filled: boolean, error?: ErrorCode): void {
    if (!filled) return
    this.filled++
    if (error) {
      this.error = this.lastError ? error : (this.error ?? error)
      return
    }
    if (n === undefined) return
    this.numbers++
    this.sum += n
    this.squares += n * n
    this.product *= n
    if (n < this.min) this.min = n
    if (n > this.max) this.max = n
  }

  /** Takes in the tally of a cell under this one. */
  merge(t: Tally): void {
    this.filled += t.filled
    this.error = this.lastError ? (t.error ?? this.error) : (this.error ?? t.error)
    if (!t.numbers) return
    this.numbers += t.numbers
    this.sum += t.sum
    this.squares += t.squares
    this.product *= t.product
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
        return n ? sum / n : { error: 'DIV0' }
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
    const v = sample ? (squares - (sum * sum) / n) / (n - 1) : (n * squares - sum * sum) / (n * n)
    const out = fn === 'stdDev' || fn === 'stdDevp' ? Math.sqrt(Math.max(v, 0)) : Math.max(v, 0)
    return Number.isFinite(out) ? out : { error: 'NUM' }
  }
}
