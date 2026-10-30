/**
 * The 11 functions of a value field over the records at a cell of the cube. A cell no value reaches (every record blank
 * in the field) is empty; an error among the values is the answer, the first in the records' order; text and booleans
 * count for `count` alone, so a sum over them is 0, an average #DIV/0!.
 */
import type { CellValue, ErrorCode, PivotFunction } from '../contract.js'
import { momentValue } from '../functions/statistical/moments.js'

/** The functions that need every number, not running totals. */
const SPREAD: ReadonlySet<PivotFunction> = new Set(['stdDev', 'stdDevp', 'var', 'varp'])

/** Whether summarizing by these functions needs the numbers themselves kept. */
export const keepsNumbers = (fns: Iterable<PivotFunction>): boolean => [...fns].some((f) => SPREAD.has(f))

/** The running state of one value field at one cell. */
export class Tally {
  /** Values not blank. */
  filled = 0
  numbers = 0
  sum = 0
  product = 1
  min = Infinity
  max = -Infinity
  error?: ErrorCode
  kept?: number[]

  constructor(keep: boolean) {
    if (keep) this.kept = []
  }

  /** A record's value: a number (a date's serial), other non-blank, an error, or blank (undefined). */
  add(n: number | undefined, filled: boolean, error?: ErrorCode): void {
    if (!filled) return
    this.filled++
    if (error) {
      this.error ??= error
      return
    }
    if (n === undefined) return
    this.numbers++
    this.sum += n
    this.product *= n
    if (n < this.min) this.min = n
    if (n > this.max) this.max = n
    this.kept?.push(n)
  }

  /** The cell's value by `fn`. */
  value(fn: PivotFunction): CellValue {
    if (!this.filled) return null
    if (this.error) return { error: this.error }
    const n = this.numbers
    switch (fn) {
      case 'sum':
        return this.sum
      case 'count':
        return this.filled
      case 'countNums':
        return n
      case 'average':
        return n ? this.sum / n : { error: 'DIV0' }
      case 'max':
        return n ? this.max : 0
      case 'min':
        return n ? this.min : 0
      case 'product':
        return n ? this.product : 0
    }
    const sample = fn === 'stdDev' || fn === 'var'
    if (n - (sample ? 1 : 0) < 1) return { error: 'DIV0' }
    const v = momentValue(this.kept!, sample ? n - 1 : n, fn === 'stdDev' || fn === 'stdDevp')
    return Number.isFinite(v) ? v : { error: 'NUM' }
  }
}
