/**
 * LINEST, LOGEST, TREND and GROWTH: least squares over one or more x variables by Householder QR (Excel
 * solves by QR too). This engine's own. LOGEST and GROWTH fit ln y. Sizes that do not fit are #REF!, then a
 * value that is not a number #VALUE!, as in Excel. A variable that depends linearly on the others (or on the
 * constant) is left out with coefficient 0, as Excel does: variables are taken by the largest norm left after
 * the ones taken (centered with a constant), a tie going to the later variable with a constant, to the earlier
 * without, as measured.
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

/**
 * y and x as observations: y one column (row), x as many rows (columns) as y has values; x defaults to 1…n.
 * One y value takes a row of x as its variables, else a column. Sizes are checked before values.
 */
function data(gy: Grid, gx: Grid | undefined, log: boolean): Data | CErr {
  const n = gy.rows * gy.cols
  const byRow = gy.rows === 1 && gy.cols > 1
  const vertical = gy.cols === 1 && (gy.rows > 1 || !gx || gx.rows === 1)
  if (gx && (vertical ? gx.rows : gx.cols) !== n) return E.REF
  const ys = values(gy)
  if (isErr(ys)) return ys
  const xs = gx ? values(gx) : ys.map((_, i) => i + 1)
  if (isErr(xs)) return xs
  // sizes first, then the logarithm's domain
  if (log && ys.some((y) => y <= 0)) return E.NUM
  const y = log ? ys.map(Math.log) : ys
  if (!gx) return { x: xs.map((x) => [x]), y, byRow }
  const k = vertical ? gx.cols : gx.rows
  const x = Array.from({ length: n }, (_, i) => Array.from({ length: k }, (_, j) => (vertical ? xs[i * k + j] : xs[j * n + i])))
  return { x, y, byRow }
}

interface Fit {
  /** Coefficients of the variables (0 for one left out), then the intercept. */
  m: number[]
  b: number
  /** (XᵀX)⁻¹ diagonal for the variables (0 for one left out; the intercept last, with a constant). */
  diag: number[]
  ssresid: number
  sstot: number
  df: number
  /** Variables in the fit. */
  used: number
}

/** A residual norm counts as 0 below this share of the column's own norm. */
const DEPENDENT = 1e-10
/** Residual norms this close count as a tie. */
const TIE = 1e-9

/** The variables Excel keeps: by the largest norm left after projecting out those kept (and the constant). */
function independent(x: number[][], constant: boolean): number[] {
  const n = x.length
  const k = x[0]?.length ?? 0
  const cols = Array.from({ length: k }, (_, j) => {
    const c = x.map((row) => row[j])
    const mean = constant ? c.reduce((s, t) => s + t, 0) / n : 0
    return c.map((t) => t - mean)
  })
  const scale = cols.map((c) => Math.hypot(...c))
  const basis: number[][] = []
  const kept: number[] = []
  let left = cols.map((_, j) => j)
  for (;;) {
    const res = left.map((j) => {
      let r = cols[j]
      for (const q of basis) {
        const d = r.reduce((s, t, i) => s + t * q[i], 0)
        r = r.map((t, i) => t - d * q[i])
      }
      return r
    })
    const norms = res.map((r) => Math.hypot(...r))
    let best = -1
    for (let i = 0; i < left.length; i++) {
      if (!(norms[i] > DEPENDENT * scale[left[i]])) continue
      const tie = best >= 0 && Math.abs(norms[i] - norms[best]) <= TIE * norms[best]
      if (best < 0 || (tie ? constant : norms[i] > norms[best])) best = i
    }
    if (best < 0) return kept.sort((a, b) => a - b)
    kept.push(left[best])
    basis.push(res[best].map((t) => t / norms[best]))
    left = left.filter((_, i) => i !== best)
  }
}

/** Least squares on the variables Excel keeps; one left out gets coefficient 0. */
function solve(d: Data, constant: boolean): Fit {
  const k = d.x[0]?.length ?? 0
  const keep = independent(d.x, constant)
  const f = qr({ x: d.x.map((row) => keep.map((j) => row[j])), y: d.y, byRow: d.byRow }, constant)
  const m = new Array<number>(k).fill(0)
  const diag = new Array<number>(k).fill(0)
  keep.forEach((j, i) => ((m[j] = f.m[i]), (diag[j] = f.diag[i])))
  return { ...f, m, diag: constant ? [...diag, f.diag[keep.length]] : diag, used: keep.length }
}

/** Least squares by Householder QR on the design matrix (a column of ones first when `constant`) of independent columns. */
function qr({ x, y }: Data, constant: boolean): Omit<Fit, 'used'> {
  const n = y.length
  const a = x.map((row) => (constant ? [1, ...row] : [...row]))
  const p = a[0].length
  const q = [...y]
  for (let j = 0; j < p; j++) {
    let norm = 0
    for (let i = j; i < n; i++) norm += a[i][j] * a[i][j]
    norm = Math.sqrt(norm)
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
  // Excel counts at least one fitted parameter: no constant and every variable left out keeps n − 1
  return { m, b: constant ? beta[0] : 0, diag: constant ? [...diag.slice(1), diag[0]] : diag, ssresid, sstot, df: n - Math.max(1, p) }
}

/** LINEST's answer: coefficients last variable first, then the intercept; with `stats` four rows more. */
function table(f: Fit, constant: boolean, stats: boolean, exp: boolean): Arr | CErr {
  const k = f.m.length
  const coef = [...f.m].reverse()
  const row1 = [...coef, f.b].map((c) => (exp ? Math.exp(c) : c))
  if (!stats) return new Arr(1, k + 1, 1, k + 1, row1)
  const w = k + 1
  const na = (xs: V[]) => [...xs, ...new Array<V>(w - xs.length).fill(E.NA)]
  // a perfect fit has 0 error, with no degrees of freedom left too
  const sey = f.ssresid === 0 ? 0 : f.df > 0 ? Math.sqrt(f.ssresid / f.df) : E.NUM
  const se = (d: number) => (typeof sey === 'number' ? sey * Math.sqrt(d) : sey)
  const ses: V[] = [...f.diag.slice(0, k)].reverse().map(se)
  ses.push(constant ? se(f.diag[k]) : E.NA)
  const ssreg = f.sstot - f.ssresid
  // Excel: r² of a constant y is 1; F counts the variables in the fit
  const r2 = f.sstot === 0 ? 1 : ssreg / f.sstot
  const F = f.df > 0 && f.ssresid > 0 && f.used > 0 ? ssreg / f.used / (f.ssresid / f.df) : E.NUM
  const v = [...row1, ...ses, ...na([r2, sey]), ...na([F, f.df]), ...na([ssreg, f.ssresid])]
  return new Arr(5, w, 5, w, v)
}

function linest(log: boolean) {
  return (gy: Grid, gx?: Grid, constant = true, stats = false): V | Arr => {
    const d = data(gy, gx, log)
    if (isErr(d)) return d
    return table(solve(d, constant), constant, stats, log)
  }
}

/** TREND and GROWTH: the fitted values at new x (by default the known x), shaped as the new x for one variable. */
function trend(log: boolean) {
  return (gy: Grid, gx?: Grid, gnew?: Grid, constant = true): V | Arr => {
    const d = data(gy, gx, log)
    if (isErr(d)) return d
    const f = solve(d, constant)
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
