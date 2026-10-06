/** MDETERM, MINVERSE, MMULT and MUNIT: matrices of numbers, by Gaussian elimination with partial pivoting. This engine's own. */
import { E, type CErr, type V } from '../../values.js'
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { Arr } from '../../arrays/value.js'

/** The rows of a grid filled with numbers only (a blank, text or a boolean is #VALUE!). */
function numbers(g: Grid): number[][] | CErr {
  const out: number[][] = []
  for (let r = 0; r < g.rows; r++) {
    const row: number[] = []
    for (let c = 0; c < g.cols; c++) {
      const v = gridAt(g, r * g.cols + c)
      if (typeof v !== 'number') return E.VALUE
      row.push(v)
    }
    out.push(row)
  }
  return out
}

const square = (g: Grid): number[][] | CErr => (g.rows === g.cols ? numbers(g) : E.VALUE)

const matrix = (m: number[][]): Arr => new Arr(m.length, m[0].length, m.length, m[0].length, m.flat())

function mdeterm(g: Grid): number | CErr {
  const a = square(g)
  if (!Array.isArray(a)) return a
  const n = a.length
  let det = 1
  for (let k = 0; k < n; k++) {
    let p = k
    for (let i = k + 1; i < n; i++) if (Math.abs(a[i][k]) > Math.abs(a[p][k])) p = i
    if (a[p][k] === 0) return 0
    if (p !== k) {
      ;[a[p], a[k]] = [a[k], a[p]]
      det = -det
    }
    det *= a[k][k]
    for (let i = k + 1; i < n; i++) {
      const f = a[i][k] / a[k][k]
      for (let j = k + 1; j < n; j++) a[i][j] -= f * a[k][j]
    }
  }
  return det
}

/** The inverse by Gauss–Jordan elimination; a singular matrix is #NUM!. */
function minverse(g: Grid): V | Arr {
  const a = square(g)
  if (!Array.isArray(a)) return a
  const n = a.length
  const inv = a.map((_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)))
  for (let k = 0; k < n; k++) {
    let p = k
    for (let i = k + 1; i < n; i++) if (Math.abs(a[i][k]) > Math.abs(a[p][k])) p = i
    if (a[p][k] === 0) return E.NUM
    ;[a[p], a[k]] = [a[k], a[p]]
    ;[inv[p], inv[k]] = [inv[k], inv[p]]
    const d = a[k][k]
    for (let j = 0; j < n; j++) {
      a[k][j] /= d
      inv[k][j] /= d
    }
    for (let i = 0; i < n; i++) {
      if (i === k || a[i][k] === 0) continue
      const f = a[i][k]
      for (let j = 0; j < n; j++) {
        a[i][j] -= f * a[k][j]
        inv[i][j] -= f * inv[k][j]
      }
    }
  }
  return matrix(inv)
}

/** The product; the columns of the first must equal the rows of the second. An error inside is the answer (kind M). */
function mmult(x: Grid, y: Grid): V | Arr {
  const a = numbers(x)
  if (!Array.isArray(a)) return a
  const b = numbers(y)
  if (!Array.isArray(b)) return b
  if (x.cols !== y.rows) return E.VALUE
  return matrix(a.map((row) => b[0].map((_, j) => row.reduce((s, v, k) => s + v * b[k][j], 0))))
}

export const functions: Record<string, FnDef> = {
  MDETERM: { args: 'M', fn: mdeterm },
  MINVERSE: { args: 'M', array: true, fn: minverse },
  MMULT: { args: 'M M', array: true, fn: mmult },
  MUNIT: { args: 'i', array: true, fn: (n: number) => (n < 1 ? E.VALUE : matrix(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))))) },
}
