/**
 * LINEST, LOGEST, TREND and GROWTH: least squares over one or more x variables by Householder QR (Excel
 * solves by QR too). This engine's own. LOGEST and GROWTH fit ln y. Sizes that do not fit are #REF!, a value
 * that is not a number #VALUE!, as in Excel.
 */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { E, isErr, type CErr, type V } from '../../values.js'

interface Data {
  /** Observations × variables. */
  x: number[][]
  y: number[]
  /** y was a row: the variables of x are its rows. */
  byRow: boolean
}

/** The numbers of a grid row by row; anything else is #VALUE!. */
function values(g: Grid): number[] | CErr {
  const out: number[] = []
  for (let i = 0; i < g.rows * g.cols; i++) {
    const v = gridAt(g, i)
    if (isErr(v)) return v
    if (typeof v !== 'number') return E.VALUE
    out.push(v)
  }
  return out
}

/** y and x as observations: y one column (row), x as many rows (columns) as y has values; x defaults to 1…n. */
function data(gy: Grid, gx: Grid | undefined, log: boolean): Data | CErr {
  const ys = values(gy)
  if (isErr(ys)) return ys
  const n = ys.length
  const byRow = gy.rows === 1 && gy.cols > 1
  const vertical = gy.cols === 1
  const xs = gx ? values(gx) : ys.map((_, i) => i + 1)
  if (isErr(xs)) return xs
  if (gx && (vertical ? gx.rows : gx.cols) !== n) return E.REF
  // sizes first, then the logarithm's domain
  if (log && ys.some((y) => y <= 0)) return E.NUM
  const y = log ? ys.map(Math.log) : ys
  if (!gx) return { x: xs.map((x) => [x]), y, byRow }
  const k = vertical ? gx.cols : gx.rows
  const x = Array.from({ length: n }, (_, i) => Array.from({ length: k }, (_, j) => (vertical ? xs[i * k + j] : xs[j * n + i])))
  return { x, y, byRow }
}

interface Fit {
  /** Coefficients of the variables, then the intercept. */
  m: number[]
  b: number
  /** (XᵀX)⁻¹ diagonal for the variables (and the intercept last, with a constant). */
  diag: number[]
  ssresid: number
  sstot: number
  df: number
}

/** Least squares by Householder QR on the design matrix (a column of ones first when `constant`). */
function solve({ x, y }: Data, constant: boolean): Fit | CErr {
  const n = y.length
  const a = x.map((row) => (constant ? [1, ...row] : [...row]))
  const p = a[0].length
  if (n < p) return E.NUM
  const q = [...y]
  for (let j = 0; j < p; j++) {
    let norm = 0
    for (let i = j; i < n; i++) norm += a[i][j] * a[i][j]
    norm = Math.sqrt(norm)
    if (norm === 0) return E.NUM
    const alpha = a[j][j] > 0 ? -norm : norm
    const v = a.map((row, i) => (i < j ? 0 : row[j]))
    v[j] -= alpha
    const vv = v.reduce((s, t) => s + t * t, 0)
    if (vv === 0) continue
    for (let c = j; c < p; c++) {
      let d = 0
      for (let i = j; i < n; i++) d += v[i] * a[i][c]
      for (let i = j; i < n; i++) a[i][c] -= (2 * d * v[i]) / vv
    }
    let d = 0
    for (let i = j; i < n; i++) d += v[i] * q[i]
    for (let i = j; i < n; i++) q[i] -= (2 * d * v[i]) / vv
  }
  // back substitution R·β = Qᵀy, and R⁻¹ for the standard errors
  const beta = new Array<number>(p).fill(0)
  for (let j = p - 1; j >= 0; j--) {
    let s = q[j]
    for (let c = j + 1; c < p; c++) s -= a[j][c] * beta[c]
    beta[j] = s / a[j][j]
  }
  const rinv = Array.from({ length: p }, () => new Array<number>(p).fill(0))
  for (let j = p - 1; j >= 0; j--) {
    rinv[j][j] = 1 / a[j][j]
    for (let i = j - 1; i >= 0; i--) {
      let s = 0
      for (let c = i + 1; c <= j; c++) s += a[i][c] * rinv[c][j]
      rinv[i][j] = -s / a[i][i]
    }
  }
  const diag = rinv.map((row) => row.reduce((s, t) => s + t * t, 0))
  let ssresid = 0
  for (let i = p; i < n; i++) ssresid += q[i] * q[i]
  const my = constant ? y.reduce((s, t) => s + t, 0) / n : 0
  const sstot = y.reduce((s, t) => s + (t - my) * (t - my), 0)
  const m = constant ? beta.slice(1) : beta
  return { m, b: constant ? beta[0] : 0, diag: constant ? [...diag.slice(1), diag[0]] : diag, ssresid, sstot, df: n - p }
}

/** LINEST's answer: coefficients last variable first, then the intercept; with `stats` four rows more. */
function table(f: Fit, constant: boolean, stats: boolean, exp: boolean): Arr | CErr {
  const k = f.m.length
  const coef = [...f.m].reverse()
  const row1 = [...coef, f.b].map((c) => (exp ? Math.exp(c) : c))
  if (!stats) return new Arr(1, k + 1, 1, k + 1, row1)
  const w = k + 1
  const na = (xs: V[]) => [...xs, ...new Array<V>(w - xs.length).fill(E.NA)]
  const sey = f.df > 0 ? Math.sqrt(f.ssresid / f.df) : E.NUM
  const se = (d: number) => (typeof sey === 'number' ? sey * Math.sqrt(d) : sey)
  const ses: V[] = [...f.diag.slice(0, k)].reverse().map(se)
  ses.push(constant ? se(f.diag[k]) : E.NA)
  const ssreg = f.sstot - f.ssresid
  const r2 = f.sstot === 0 ? E.NUM : ssreg / f.sstot
  const F = f.df > 0 && f.ssresid > 0 ? ssreg / k / (f.ssresid / f.df) : E.NUM
  const v = [...row1, ...ses, ...na([r2, sey]), ...na([F, f.df]), ...na([ssreg, f.ssresid])]
  return new Arr(5, w, 5, w, v)
}

function linest(log: boolean) {
  return (gy: Grid, gx?: Grid, constant = true, stats = false): V | Arr => {
    const d = data(gy, gx, log)
    if (isErr(d)) return d
    const f = solve(d, constant)
    return isErr(f) ? f : table(f, constant, stats, log)
  }
}

/** TREND and GROWTH: the fitted values at new x (by default the known x), shaped as the new x for one variable. */
function trend(log: boolean) {
  return (gy: Grid, gx?: Grid, gnew?: Grid, constant = true): V | Arr => {
    const d = data(gy, gx, log)
    if (isErr(d)) return d
    const f = solve(d, constant)
    if (isErr(f)) return f
    const k = f.m.length
    let points: number[][]
    let rows: number
    let cols: number
    if (!gnew) {
      points = d.x
      ;[rows, cols] = d.byRow ? [1, d.y.length] : [d.y.length, 1]
    } else {
      const xs = values(gnew)
      if (isErr(xs)) return xs
      if (k === 1) {
        points = xs.map((x) => [x])
        ;[rows, cols] = [gnew.rows, gnew.cols]
      } else {
        const vertical = !d.byRow
        if ((vertical ? gnew.cols : gnew.rows) !== k) return E.REF
        const n = xs.length / k
        points = Array.from({ length: n }, (_, i) => Array.from({ length: k }, (_, j) => (vertical ? xs[i * k + j] : xs[j * n + i])))
        ;[rows, cols] = vertical ? [n, 1] : [1, n]
      }
    }
    const out = points.map((p) => {
      const z = f.b + p.reduce((s, x, j) => s + x * f.m[j], 0)
      return log ? Math.exp(z) : z
    })
    return new Arr(rows, cols, rows, cols, out)
  }
}

export const functions: Record<string, FnDef> = {
  GROWTH: { args: 'G G? G? b?', array: true, fn: trend(true) },
  LINEST: { args: 'G G? b? b?', array: true, fn: linest(false) },
  LOGEST: { args: 'G G? b? b?', array: true, fn: linest(true) },
  TREND: { args: 'G G? G? b?', array: true, fn: trend(false) },
}
