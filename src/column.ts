/**
 * A column of a sheet's values by row, in chunks of 256 rows, each made when one of its rows is first written. A chunk of
 * fewer than `FEW` values keeps them as pairs of a row and its value, read by a short scan; at `FEW` it becomes an array
 * by row, read by index. A JavaScript array with a value far below the rest is kept by the engine as a dictionary, read
 * by hashing, and a dense chunk is never sparser than its own rows; a sparse sheet holds its cells, not empty rows.
 */
import type { V } from './values.js'

const BITS = 8
const MASK = (1 << BITS) - 1
/** A chunk with this many values becomes an array by row. */
const FEW = 20

/** The values of a sparse chunk: rows in order and their values. */
interface Pairs {
  rows: number[]
  values: V[]
}

export class Column {
  private readonly dense: (V[] | undefined)[] = []
  private readonly sparse: (Pairs | undefined)[] = []
  /** One past the last row written, as an array's length. */
  length = 0

  /** The value of row `r`; undefined where none was written. */
  get(r: number): V | undefined {
    const k = r >> BITS
    const d = this.dense[k]
    if (d !== undefined) return d[r & MASK]
    const p = this.sparse[k]
    if (p === undefined) return undefined
    const j = p.rows.indexOf(r & MASK)
    return j < 0 ? undefined : p.values[j]
  }

  /** Whether row `r` was written (a null too). */
  has(r: number): boolean {
    const k = r >> BITS
    const d = this.dense[k]
    if (d !== undefined) return (r & MASK) in d
    return this.sparse[k]?.rows.includes(r & MASK) ?? false
  }

  set(r: number, v: V): void {
    const k = r >> BITS, i = r & MASK
    const d = this.dense[k]
    if (d !== undefined) d[i] = v
    else {
      const p = this.sparse[k]
      if (p === undefined) this.sparse[k] = { rows: [i], values: [v] }
      else {
        const j = p.rows.indexOf(i)
        if (j >= 0) p.values[j] = v
        else if (p.rows.length + 1 < FEW) {
          let at = p.rows.length
          while (at > 0 && p.rows[at - 1] > i) at--
          p.rows.splice(at, 0, i)
          p.values.splice(at, 0, v)
        } else {
          const out: V[] = []
          p.rows.forEach((row, n) => (out[row] = p.values[n]))
          out[i] = v
          this.dense[k] = out
          this.sparse[k] = undefined
        }
      }
    }
    if (r >= this.length) this.length = r + 1
  }

  /** Each row written, in order, with its value. */
  forEach(f: (v: V, r: number) => void): void {
    const n = Math.max(this.dense.length, this.sparse.length)
    for (let k = 0; k < n; k++) {
      const d = this.dense[k]
      if (d !== undefined) d.forEach((v, i) => f(v, (k << BITS) | i))
      else {
        const p = this.sparse[k]
        if (p) p.rows.forEach((row, j) => f(p.values[j], (k << BITS) | row))
      }
    }
  }
}
