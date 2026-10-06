/** Rates found by iteration (RATE, IRR, XIRR) and the cash-flow functions next to them (MIRR, XNPV), this engine's own. */
import { gridAt, type Cx, type FnDef, type Grid } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { atp } from '../date/atp.js'

/** Newton's method on a rate: null when it does not settle (|Δ| < 1e-7) within `max` steps or settles at −1 or below. */
function newton(f: (x: number) => [number, number], guess: number, max: number): number | null {
  let x = guess
  for (let k = 0; k < max; k++) {
    const [y, dy] = f(x)
    if (Math.abs(y) < 1e-14) return x > -1 ? x : null
    const next = dy === 0 ? x + 1.1e-7 : x - y / dy
    if (!Number.isFinite(next)) return null
    if (Math.abs(next - x) < 1e-7) return next > -1 ? next : null
    x = next
  }
  return null
}

/** RATE's balance fv + pv·(1+r)^n + pmt·((1+r)^n−1)/r, payments at the start folded into pv and fv. */
function annuity(nper: number, pmt: number, pv: number, fv: number, type: number) {
  if (type) [pv, fv] = [pv + pmt, fv - pmt]
  return (r: number): [number, number] => {
    const g1 = Math.pow(1 + r, nper - 1)
    const g = g1 * (1 + r)
    const s = r === 0 ? nper : (g - 1) / r
    const ds = r === 0 ? (nper * (nper - 1)) / 2 : (nper * g1) / r - s / r
    return [fv + pv * g + pmt * s, pv * nper * g1 + pmt * ds]
  }
}

/** IRR by Newton's method in v = 1/(1+r), where the present value is a polynomial (Excel converges from far guesses too). */
function irr(xs: number[], guess: number): number | null {
  let v = 1 / (1 + guess)
  for (let k = 0; k < 20; k++) {
    let y = 0
    let dy = 0
    for (let i = xs.length - 1; i >= 0; i--) {
      dy = dy * v + y
      y = y * v + xs[i]
    }
    const next = v - y / dy
    if (!Number.isFinite(next) || next === 0) return null
    if (Math.abs(1 / next - 1 / v) < 1e-7) return next > 0 ? 1 / next - 1 : null
    v = next
  }
  return null
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

/**
 * XNPV: a blank is #VALUE!, text or a boolean #NUM!; XIRR (`lenient`, the book's rules) reads numeric and date
 * text and a blank value as 0, a boolean or other text is #VALUE! anywhere before a blank date counts as #NUM!.
 */
function dated(values: Grid, dates: Grid, lenient: Rules | null): [number[], number[]] | CErr {
  const n = values.rows * values.cols
  if (n !== dates.rows * dates.cols) return E.NUM
  const vs: number[] = []
  const ds: number[] = []
  let blank = false
  for (let i = 0; i < n; i++) {
    let v: V = gridAt(values, i)
    let d: V = gridAt(dates, i)
    if (isErr(v)) return v
    if (isErr(d)) return d
    if (lenient) {
      if (typeof v === 'boolean' || typeof d === 'boolean') return E.VALUE
      v = lenient.toNum(v)
      d = typeof d === 'string' ? (lenient.dateText(d) ?? lenient.toNum(d)) : d
      if (isErr(v)) return v
      if (isErr(d)) return d
      if (d === null) {
        blank = true
        continue
      }
    }
    if (v === null || d === null) return E.VALUE
    if (typeof v !== 'number' || typeof d !== 'number' || d < 0) return E.NUM
    vs.push(v)
    ds.push(Math.floor(d))
  }
  return blank ? E.NUM : [vs, ds]
}

const xnpv = (rate: number, vs: number[], ds: number[]) => vs.reduce((s, v, i) => s + v / Math.pow(1 + rate, (ds[i] - ds[0]) / 365), 0)

/** XIRR's root: Newton's method from the guess, bisection over (−1, 1e10) when it fails. */
function xirr(vs: number[], ds: number[], guess: number): number | null {
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
  let x = guess
  for (let k = 0; k < 100; k++) {
    const [y, dy] = f(x)
    const next = x - y / dy
    if (!(next > -1) || !Number.isFinite(next)) break
    if (Math.abs(next - x) < 1e-12 * Math.max(1, Math.abs(next))) return next
    x = next
  }
  let a = -1 + 1e-12
  let b = Math.max(1, guess)
  const fa = f(a)[0]
  while (Math.sign(f(b)[0]) === Math.sign(fa) && b < 1e10) b *= 2
  if (Math.sign(f(b)[0]) === Math.sign(fa)) return null
  for (let k = 0; k < 400 && b - a > 1e-15 * Math.max(1, Math.abs(a)); k++) {
    const m = (a + b) / 2
    if (Math.sign(f(m)[0]) === Math.sign(fa)) a = m
    else b = m
  }
  return (a + b) / 2
}

export const functions: Record<string, FnDef> = {
  // at most 20 steps, as Excel documents
  IRR: {
    args: 'G n?',
    fn: (g: Grid, guess = 0.1) => {
      const xs = flows(g)
      if (isErr(xs)) return xs
      if (!signs(xs)) return E.NUM
      return irr(xs, guess) ?? E.NUM
    },
  },
  MIRR: {
    args: 'N n n',
    fn: (xs: number[], finance: number, reinvest: number) => {
      if (reinvest === -1) return E.DIV0
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
      nper <= 0 ? E.NUM : (newton(annuity(nper, pmt, pv, fv, type), guess, 150) ?? E.NUM),
  },
  // Excel's answers lie within its 1e-8 tolerance of the root; from a negative guess it finds no positive rate
  XIRR: atp('G G x?', (cx: Cx, values: Grid, dates: Grid, guess = 0.1) => {
    const vd = dated(values, dates, cx.rules)
    if (isErr(vd)) return vd
    const [vs, ds] = vd
    if (!signs(vs)) return E.NUM
    const r = xirr(vs, ds, guess)
    return r === null || (guess < 0 && r >= 0) ? E.NUM : r
  }, { cx: true }),
  // the rate must be positive and no date may precede the first
  XNPV: atp('x G G', (rate: number, values: Grid, dates: Grid) => {
    if (rate <= 0) return E.NUM
    const vd = dated(values, dates, null)
    if (isErr(vd)) return vd
    const [vs, ds] = vd
    if (ds.some((d) => d < ds[0])) return E.NUM
    return xnpv(rate, vs, ds)
  }),
}
