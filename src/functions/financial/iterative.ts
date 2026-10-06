/** Rates found by iteration (RATE, IRR, XIRR) and the cash-flow functions next to them (MIRR, XNPV), this engine's own. */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { E, isErr, toNum, type CErr, type V } from '../../values.js'
import { atp } from '../date/atp.js'

/** Newton's method from `guess`; null when it leaves (−1, ∞) or does not settle within `max` steps. */
function newton(f: (x: number) => [number, number], guess: number, max: number, eps: number): number | null {
  let x = guess
  for (let k = 0; k < max; k++) {
    const [y, dy] = f(x)
    if (!Number.isFinite(y) || !Number.isFinite(dy) || dy === 0) return null
    const next = x - y / dy
    if (!(next > -1)) return null
    if (Math.abs(next - x) < eps) return next
    x = next
  }
  return null
}

/** Annuity balance at rate r and its derivative: pv·(1+r)^n + pmt·(1+r·t)·((1+r)^n−1)/r + fv. */
function annuity(nper: number, pmt: number, pv: number, fv: number, t: number) {
  return (r: number): [number, number] => {
    if (r === 0) return [pv + pmt * nper + fv, pv * nper + pmt * (t * nper + (nper * (nper - 1)) / 2)]
    const g = Math.pow(1 + r, nper)
    const dg = nper * Math.pow(1 + r, nper - 1)
    const s = (g - 1) / r
    const ds = (dg * r - (g - 1)) / (r * r)
    return [pv * g + pmt * (1 + r * t) * s + fv, pv * dg + pmt * t * s + pmt * (1 + r * t) * ds]
  }
}

/** Numbers of a range in order (IRR: text, booleans and blanks skipped, an error is #VALUE!). */
function flows(g: Grid): number[] | CErr {
  const out: number[] = []
  for (const v of g.v) {
    if (isErr(v)) return E.VALUE
    if (typeof v === 'number') out.push(v)
  }
  return out
}

const signs = (xs: number[]) => xs.some((x) => x > 0) && xs.some((x) => x < 0)

/** XNPV and XIRR: values and dates of one size; XNPV takes numbers only, XIRR reads text numbers and blanks as 0. */
function dated(values: Grid, dates: Grid, lenient: boolean): [number[], number[]] | CErr {
  const n = values.rows * values.cols
  if (n !== dates.rows * dates.cols) return E.NUM
  const vs: number[] = []
  const ds: number[] = []
  for (let i = 0; i < n; i++) {
    let v: V = gridAt(values, i)
    const d = gridAt(dates, i)
    if (isErr(v)) return v
    if (isErr(d)) return d
    if (lenient && typeof v !== 'number' && typeof v !== 'boolean') v = toNum(v)
    if (isErr(v)) return v
    if (typeof v !== 'number') return E.NUM
    if (typeof d !== 'number') return E.VALUE
    if (d < 0) return E.NUM
    vs.push(v)
    ds.push(Math.floor(d))
  }
  return [vs, ds]
}

const xnpv = (rate: number, vs: number[], ds: number[]) => vs.reduce((s, v, i) => s + v / Math.pow(1 + rate, (ds[i] - ds[0]) / 365), 0)

export const functions: Record<string, FnDef> = {
  // Newton's method from 0.1, at most 20 steps, as Excel documents
  IRR: {
    args: 'G n?',
    fn: (g: Grid, guess = 0.1) => {
      const xs = flows(g)
      if (isErr(xs)) return xs
      if (!signs(xs)) return E.NUM
      const f = (r: number): [number, number] => {
        let y = 0
        let dy = 0
        for (let i = 0; i < xs.length; i++) {
          y += xs[i] / Math.pow(1 + r, i)
          dy -= (i * xs[i]) / Math.pow(1 + r, i + 1)
        }
        return [y, dy]
      }
      return newton(f, guess, 20, 1e-7) ?? E.NUM
    },
  },
  MIRR: {
    args: 'N n n',
    fn: (xs: number[], finance: number, reinvest: number) => {
      const n = xs.length
      let pos = 0
      let neg = 0
      for (let i = 0; i < n; i++) {
        if (xs[i] > 0) pos += xs[i] * Math.pow(1 + reinvest, n - 1 - i)
        else neg += xs[i] / Math.pow(1 + finance, i)
      }
      if (pos === 0 || neg === 0) return E.DIV0
      return Math.pow(-pos / neg, 1 / (n - 1)) - 1
    },
  },
  RATE: {
    args: 'n n n n? n? n?',
    fn: (nper: number, pmt: number, pv: number, fv = 0, type = 0, guess = 0.1) =>
      nper <= 0 ? E.NUM : (newton(annuity(nper, pmt, pv, fv, type ? 1 : 0), guess, 20, 1e-7) ?? E.NUM),
  },
  XIRR: atp('G G x?', (values: Grid, dates: Grid, guess = 0.1) => {
    const vd = dated(values, dates, true)
    if (isErr(vd)) return vd
    const [vs, ds] = vd
    if (!signs(vs)) return E.NUM
    const f = (r: number): [number, number] => {
      let y = 0
      let dy = 0
      for (let i = 0; i < vs.length; i++) {
        const t = (ds[i] - ds[0]) / 365
        y += vs[i] / Math.pow(1 + r, t)
        dy -= (t * vs[i]) / Math.pow(1 + r, t + 1)
      }
      return [y, dy]
    }
    return newton(f, guess, 100, 1e-8) ?? E.NUM
  }),
  // the rate must be positive and no date may precede the first
  XNPV: atp('x G G', (rate: number, values: Grid, dates: Grid) => {
    if (rate <= 0) return E.NUM
    const vd = dated(values, dates, false)
    if (isErr(vd)) return vd
    const [vs, ds] = vd
    if (ds.some((d) => d < ds[0])) return E.NUM
    return xnpv(rate, vs, ds)
  }),
}
