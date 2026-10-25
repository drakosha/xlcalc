/** MDETERM, MINVERSE, MMULT and MUNIT: matrices of numbers, by Gaussian elimination with partial pivoting. This engine's own. */
import { E, isErr, type CErr, type V } from '../../values.js'
import { observations, values, type FnDef, type Grid } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { allot, spend } from '../../exec/budget.js'

/** The rows of a grid of numbers (a blank, text or a boolean is #VALUE!; kind M answers an error before it is read). */
function numbers(g: Grid): number[][] | CErr {
  const xs = values(g)
  return isErr(xs) ? xs : observations(xs, g.rows, g.cols, true)
}

const square = (g: Grid): number[][] | CErr => (g.rows === g.cols ? numbers(g) : E.VALUE)

const matrix = (m: number[][]): Arr => new Arr(m.length, m[0].length, m.length, m[0].length, m.flat())

const identity = (n: number) => Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)))

/** Partial pivoting: the row from `k` down whose value in column `k` is the largest in magnitude. */
function pivot(a: number[][], k: number): number {
  let p = k
  for (let i = k + 1; i < a.length; i++) if (Math.abs(a[i][k]) > Math.abs(a[p][k])) p = i
  return p
}

function mdeterm(g: Grid): number | CErr {
  const a = square(g)
  if (!Array.isArray(a)) return a
  const n = a.length
  let det = 1
  for (let k = 0; k < n; k++) {
    spend(n * n)
    const p = pivot(a, k)
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
  const inv = identity(n)
  for (let k = 0; k < n; k++) {
    spend(2 * n * n)
    const p = pivot(a, k)
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
  allot(x.rows * y.cols)
  return matrix(
    a.map((row) => {
      spend(y.cols * x.cols)
      return b[0].map((_, j) => row.reduce((s, v, k) => s + v * b[k][j], 0))
    }),
  )
}

export const functions: Record<string, FnDef> = {
  MDETERM: { args: 'M', fn: mdeterm },
  MINVERSE: { args: 'M', array: true, fn: minverse },
  MMULT: { args: 'M M', array: true, fn: mmult },
  MUNIT: {
    args: 'i',
    array: true,
    fn: (n: number) => {
      if (n < 1) return E.VALUE
      allot(n * n)
      return matrix(identity(n))
    },
  },
}
