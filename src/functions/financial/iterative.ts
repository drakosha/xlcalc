/** Rates found by iteration (RATE, IRR, XIRR) and the cash-flow functions next to them (MIRR, XNPV), this engine's own. */
import { gridAt, type Arg, type Cx, type FnDef, type Grid } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, MIN_NORMAL, type CErr, type V } from '../../values.js'
import { atp } from '../date/atp.js'
import { areaNumbers, numbersOf } from '../statistical/common.js'

/**
 * Newton's method on a rate: null when it does not settle (|Δ| < 1e-7) within `max` steps or settles at −1 or
 * below. A balance below 1e-14 settles only with a small step too: one that vanishes only as the rate goes to −1
 * (`RATE(48,0,8000)`) does not settle, as in Excel.
 */
function newton(f: (x: number) => [number, number], guess: number, max: number): number | null {
  let x = guess
  for (let k = 0; k < max; k++) {
    const [y, dy] = f(x)
    if (Math.abs(y) < 1e-14 && Math.abs(y) < 1e-7 * Math.abs(dy)) return x > -1 ? x : null
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

/** Excel's IRR evaluates the cash flows in order at s = 1/(1+r) − 1. */
function irrBalance(xs: number[], first: number, s: number): number | CErr {
  if (!Number.isFinite(s)) return E.NUM
  const factor = 1 / (1 + s)
  let discount = 1
  let balance = xs[first]
  for (let i = first + 1; i < xs.length; i++) {
    discount *= factor
    if (!Number.isFinite(discount) || Math.abs(discount) < MIN_NORMAL) return E.DIV0
    balance += xs[i] / discount
  }
  // A later invalid divisor takes precedence over an overflow in the sum.
  return Number.isFinite(balance) ? balance : E.NUM
}

function irrAttempt(xs: number[], first: number, guess: number): number | CErr | null {
  const step = 0.001
  const tolerance = 1e-7
  let previousRate = guess / (-1 - guess)
  let previousBalance = irrBalance(xs, first, previousRate)
  if (isErr(previousBalance)) return previousBalance
  let rate = previousRate + (previousBalance < 0 ? -step : step)
  let balance = irrBalance(xs, first, rate)
  if (isErr(balance)) return balance

  for (let i = 0; i < 200; i++) {
    if (balance === previousBalance) {
      // Equal balances give the secant a zero denominator, so reevaluate the older point.
      previousRate += previousRate > rate ? -step : step
      previousBalance = irrBalance(xs, first, previousRate)
      if (isErr(previousBalance)) return previousBalance
      if (balance === previousBalance) return null
    }
    const change = ((rate - previousRate) * balance) / (balance - previousBalance)
    previousRate = rate
    previousBalance = balance
    rate -= change
    balance = irrBalance(xs, first, rate)
    if (isErr(balance)) return balance
    if (Math.abs(balance) <= tolerance && Math.abs(rate - previousRate) <= tolerance) {
      return rate - tolerance > -1 ? 1 / (1 + rate) - 1 : null
    }
  }

  const finalTolerance = 1e-4
  if (Math.abs(balance) <= finalTolerance && Math.abs(rate - previousRate) <= finalTolerance && rate - finalTolerance > -1) {
    return 1 / (1 + rate) - 1
  }
  return null
}

const IRR_DEFAULT_GUESS = 0.1
const IRR_RESTARTS = [IRR_DEFAULT_GUESS, -1 / 11, 2]

/** Excel retries an unsettled secant search from three fixed guesses. Arithmetic errors end the search. */
function irr(xs: number[], guess: number): number | CErr {
  // Excel starts discounting after the first nonzero cash flow.
  const first = xs.findIndex((value) => value !== 0)
  const guesses = guess === IRR_DEFAULT_GUESS ? IRR_RESTARTS : [guess, ...IRR_RESTARTS]
  for (const start of guesses) {
    const result = irrAttempt(xs, first, start)
    if (result !== null) return result
  }
  return E.NUM
}

/** Numbers of a range in order (IRR: text, booleans and blanks skipped, an error is #VALUE!, where MIRR answers it). */
function flows(g: Grid): number[] | CErr {
  const xs = numbersOf(g.v)
  return isErr(xs) ? E.VALUE : xs
}

const signs = (xs: number[]) => xs.some((x) => x > 0) && xs.some((x) => x < 0)

/**
 * XNPV: a blank first value or date is #VALUE!, a later one #NUM!, text or a boolean #NUM!; XIRR (`lenient`, the book's rules) reads numeric and date text and a blank value as 0, a boolean or
 * other text is #VALUE! anywhere before a blank date counts as #NUM!. One
 * value as the values or the dates is read as a number first: a boolean or text that is no number is #VALUE!, before
 * the sizes are compared (`XIRR(TRUE,A1:A5)`, the set `single-values`).
 */
function dated(values: Grid, dates: Grid, rules: Rules, lenient: boolean): [number[], number[]] | CErr {
  for (const g of [values, dates]) {
    const v = g.rows * g.cols === 1 ? g.v[0] : null
    if (typeof v === 'boolean' || (typeof v === 'string' && isErr(rules.toNum(v)))) return E.VALUE
  }
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
      v = rules.toNum(v)
      d = typeof d === 'string' ? (rules.dateText(d) ?? rules.toNum(d)) : d
      if (isErr(v)) return v
      if (isErr(d)) return d
      if (d === null) {
        blank = true
        continue
      }
    }
    // XNPV: a blank first value or date is #VALUE!, a later one #NUM!
    if (v === null || d === null) {
      if (i === 0) return E.VALUE
      blank = true
      continue
    }
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
  IRR: {
    args: 'G n?',
    fn: (g: Grid, guess = IRR_DEFAULT_GUESS) => {
      const xs = flows(g)
      if (isErr(xs)) return xs
      // a guess of -1 or less is #VALUE!, before the flows' signs are looked at
      if (guess <= -1) return E.VALUE
      if (!signs(xs)) return E.NUM
      return irr(xs, guess)
    },
  },
  // the values' numbers in order, each area of a union; one value given is an array of it, text there no number (`MIRR("x",0.1,0.1)` is #DIV/0!)
  MIRR: {
    args: 'R n n',
    cx: true,
    areas: true,
    fn: (cx: Cx, values: Arg, finance: number, reinvest: number) => {
      const xs = areaNumbers(cx, values)
      if (isErr(xs)) return xs
      if (reinvest === -1) return E.DIV0
      const n = xs.length
      let pos = 0
      let neg = 0
      for (let i = 0; i < n; i++) {
        if (xs[i] > 0) pos += xs[i] * Math.pow(1 + reinvest, n - 1 - i)
        else neg += xs[i] / Math.pow(1 + finance, i)
      }
      // no outflow, or one value alone, is #DIV/0!; no inflow -1, as in Excel
      if (neg === 0 || n < 2) return E.DIV0
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
    const vd = dated(values, dates, cx.rules, true)
    if (isErr(vd)) return vd
    const [vs, ds] = vd
    if (!signs(vs)) return E.NUM
    const r = xirr(vs, ds, guess)
    return r === null || (guess < 0 && r >= 0) ? E.NUM : r
  }, { cx: true }),
  // the rate must be positive, no date may precede the first, and a date of 0 is #NUM! (XIRR takes one)
  XNPV: atp('x G G', (cx: Cx, rate: number, values: Grid, dates: Grid) => {
    if (rate <= 0) return E.NUM
    const vd = dated(values, dates, cx.rules, false)
    if (isErr(vd)) return vd
    const [vs, ds] = vd
    if (ds.some((d) => d < ds[0] || d < 1)) return E.NUM
    return xnpv(rate, vs, ds)
  }, { cx: true }),
}
