/**
 * A column of a sheet's values by row, in chunks of 1,024 rows, each made when one of its rows is first written. A
 * JavaScript array with a value far below the rest is kept by the engine as a dictionary, and each read of it is a lookup
 * by hash; a chunk is never sparser than its own rows, so reads stay by index, and a sheet with a cell in its last row
 * holds one chunk for it, not a million empty rows.
 */
import type { V } from './values.js'

const BITS = 10
const MASK = (1 << BITS) - 1

export class Column {
  private readonly chunks: (V[] | undefined)[] = []
  /** One past the last row written, as an array's length. */
  length = 0

  /** The value of row `r`; undefined where none was written. */
  get(r: number): V | undefined {
    return this.chunks[r >> BITS]?.[r & MASK]
  }

  /** Whether row `r` was written (a null too). */
  has(r: number): boolean {
    const chunk = this.chunks[r >> BITS]
    return chunk !== undefined && (r & MASK) in chunk
  }

  set(r: number, v: V): void {
    ;(this.chunks[r >> BITS] ??= [])[r & MASK] = v
    if (r >= this.length) this.length = r + 1
  }

  /** Each row written, in order, with its value. */
  forEach(f: (v: V, r: number) => void): void {
    this.chunks.forEach((chunk, k) => chunk!.forEach((v, i) => f(v, (k << BITS) | i)))
  }
}
