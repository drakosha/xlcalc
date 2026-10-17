/**
 * Annuities, interest rates and dollar fractions.
 * Ported from formula.js 4.6.1 (src/financial.js), Copyright (c) 2014 Sutoiku, Inc., MIT; fixed against Excel.
 */
import { E, isErr, type CErr } from '../../values.js'
import { coerce, type Arg, type Cx, type FnDef } from '../../args.js'
import { atp } from '../date/atp.js'
import { power } from '../math/arithmetic.js'

function fv(rate: number, nper: number, pmt: number, pv: number, type: number): number | CErr {
  if (rate === 0) return -(pv + pmt * nper)
  // (1 + rate)^nper as Excel's power takes it: at a rate of -1, 0^0 is #NUM! and 0 to a negative power #DIV/0!
  const term = power(1 + rate, nper)
  if (isErr(term)) return term
  return -(pv * term + (pmt * (1 + rate * type) * (term - 1)) / rate)
}

/** Excel's PMT is #NUM! for a rate of −1 or less (FV and PV are not). */
function pmt(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate <= -1) return NaN
  if (rate === 0) return -(pv + fv) / nper
  const term = Math.pow(1 + rate, nper)
  return (-(pv * term + fv) * rate) / ((1 + rate * type) * (term - 1))
}

/** Interest and principal of payment `per` (a fractional `per` taken as is); the principal grows geometrically, which keeps it exact. */
function parts(rate: number, per: number, nper: number, pv: number, fvalue: number, type: number): [number, number] | CErr {
  if (per < 1 || per >= nper + 1) return E.NUM
  const p = pmt(rate, nper, pv, fvalue, type)
  if (!Number.isFinite(p)) return E.NUM
  if (rate === 0 || (type && per === 1)) return [0, p]
  const principal = ((-(pv + fvalue) * rate) / (Math.pow(1 + rate, nper) - 1)) * Math.pow(1 + rate, per - 1 - type)
  return [p - principal, principal]
}

/** Any non-zero type is payment at the beginning of the period. */
const t01 = (type: number | undefined) => (type ? 1 : 0)

/** The smallest power of 10 not below `f`: the digits DOLLARDE and DOLLARFR shift by. */
function tens(f: number): number {
  let p = 1
  while (p < f) p *= 10
  return p
}

/** CUMIPMT and CUMPRINC: rate, nper and pv positive, 1 ≤ start ≤ end ≤ nper, type exactly 0 or 1; whole periods ⌈start⌉…⌊end⌋, the principal summed as a geometric series. */
function cumulative(k: 0 | 1) {
  return atp('x x x x x x', (rate: number, nper: number, pv: number, start: number, end: number, type: number) => {
    if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper || (type !== 0 && type !== 1)) return E.NUM
    const a = Math.ceil(start)
    const b = Math.floor(end)
    if (a > b) return 0
    const p = pmt(rate, nper, pv, 0, type)
    // with payments in advance the first payment is all principal and the series starts at the second
    const first = type && a === 1 ? p : 0
    const from = a + (first ? 1 : 0)
    const unit = (-pv * rate) / (Math.pow(1 + rate, nper) - 1)
    const principal = first + (from > b ? 0 : (unit * Math.pow(1 + rate, from - 1 - type) * (Math.pow(1 + rate, b - from + 1) - 1)) / rate)
    return k ? principal : (b - a + 1) * p - principal
  })
}

const part = (k: 0 | 1) => (rate: number, per: number, nper: number, pv: number, fvalue: number, type: number): number | CErr => {
  const x = parts(rate, per, nper, pv, fvalue, type)
  return isErr(x) ? x : x[k]
}
const ipmt = part(0)
const ppmt = part(1)

export const functions: Record<string, FnDef> = {
  CUMIPMT: cumulative(0),
  CUMPRINC: cumulative(1),
  DOLLARDE: atp('x j', (x: number, f: number) => {
    if (f < 0) return E.NUM
    if (f < 1) return E.DIV0
    const int = Math.trunc(x)
    return int + ((x - int) * tens(f)) / f
  }),
  DOLLARFR: atp('x j', (x: number, f: number) => {
    if (f < 0) return E.NUM
    if (f < 1) return E.DIV0
    const int = Math.trunc(x)
    return int + ((x - int) * f) / tens(f)
  }),
  EFFECT: atp('x j', (rate: number, n: number) => (rate <= 0 || n < 1 ? E.NUM : Math.pow(1 + rate / n, n) - 1)),
  FV: { args: 'n n n n? n?', fn: (rate: number, nper: number, p: number, pv = 0, type?: number) => fv(rate, nper, p, pv, t01(type)) },
  // blanks in the schedule are 0, text and booleans #VALUE!; the principal is read first
  FVSCHEDULE: {
    args: 'x R',
    cx: true,
    fn: (cx: Cx, principal: number, schedule: Arg) => {
      const rates = coerce('Z', schedule, cx) as number[] | CErr
      return isErr(rates) ? rates : rates.reduce((acc, r) => acc * (1 + r), principal)
    },
  },
  IPMT: { args: 'n n n n n? n?', fn: (rate: number, per: number, nper: number, pv: number, fvalue = 0, type?: number) => ipmt(rate, per, nper, pv, fvalue, t01(type)) },
  // the period and the count of periods are not truncated: ISPMT(0.1,1.5,4,1000) is -62.5
  ISPMT: { args: 'n n n n', fn: (rate: number, per: number, nper: number, pv: number) => (nper === 0 ? E.DIV0 : pv * rate * (per / nper - 1)) },
  NOMINAL: atp('x j', (rate: number, n: number) => (rate <= 0 || n < 1 ? E.NUM : (Math.pow(1 + rate, 1 / n) - 1) * n)),
  NPER: {
    args: 'n n n n? n?',
    fn: (rate: number, p: number, pv: number, fvalue = 0, type?: number) => {
      if (rate === 0) return p === 0 ? E.DIV0 : -(pv + fvalue) / p
      if (rate <= -1) return E.NUM
      const t = t01(type)
      return Math.log((p * (1 + rate * t) - fvalue * rate) / (pv * rate + p * (1 + rate * t))) / Math.log(1 + rate)
    },
  },
  NPV: {
    args: 'n N',
    rep: 1,
    fn: (rate: number, values: number[]) => {
      if (rate === -1) return E.DIV0
      let v = 0
      for (let j = 0; j < values.length; j++) v += values[j] / Math.pow(1 + rate, j + 1)
      return v
    },
  },
  PDURATION: {
    args: 'n n n',
    fn: (rate: number, pv: number, fv: number) => (rate <= 0 || pv <= 0 || fv <= 0 ? E.NUM : (Math.log(fv) - Math.log(pv)) / Math.log(1 + rate)),
  },
  PMT: { args: 'n n n n? n?', fn: (rate: number, nper: number, pv: number, fvalue = 0, type?: number) => pmt(rate, nper, pv, fvalue, t01(type)) },
  PPMT: { args: 'n n n n n? n?', fn: (rate: number, per: number, nper: number, pv: number, fvalue = 0, type?: number) => ppmt(rate, per, nper, pv, fvalue, t01(type)) },
  PV: {
    args: 'n n n n? n?',
    fn: (rate: number, nper: number, p: number, fvalue = 0, type?: number) => {
      if (rate === 0) return -(fvalue + p * nper)
      // (1 + rate)^nper as in FV, then a term of 0 divides by 0
      const term = power(1 + rate, nper)
      if (isErr(term)) return term
      return term === 0 ? E.DIV0 : -(fvalue + (p * (1 + rate * t01(type)) * (term - 1)) / rate) / term
    },
  },
  RRI: { args: 'n n n', fn: (nper: number, pv: number, fv: number) => (nper <= 0 || pv === 0 ? E.NUM : Math.pow(fv / pv, 1 / nper) - 1) },
}
