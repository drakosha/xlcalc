/** Securities without periodic coupons: discounted, interest at maturity, Treasury bills. */
import type { FnDef } from '../../args.js'
import { addMonths, daysInMonth, serialOf, ymd } from '../../dates.js'
import { E } from '../../values.js'
import { atp } from '../date/atp.js'
import { validBasis, yearFrac } from '../date/daycount.js'

/** settlement < maturity, a valid basis and positive amounts, as the ToolPak checks them. */
const bad = (settlement: number, maturity: number, basis: number, ...positive: number[]) =>
  settlement >= maturity || !validBasis(basis) || positive.some((x) => x <= 0)

/** Securities of (settlement, maturity, a, b, basis) on the year fraction between the dates. */
function onYears(answer: (years: number, a: number, b: number) => number) {
  return atp('e e x x j?', (settlement: number, maturity: number, a: number, b: number, basis = 0) =>
    bad(settlement, maturity, basis, a, b) ? E.NUM : answer(yearFrac(settlement, maturity, basis), a, b),
  )
}

/** Treasury bills: settlement < maturity ≤ settlement plus one year, the third argument positive; days are actual. */
function tbill(answer: (dsm: number, x: number) => number) {
  return atp('e e x', (settlement: number, maturity: number, x: number) => {
    const s = ymd(settlement)
    const t = addMonths(s.y, s.m, 12)
    const yearOn = serialOf(t.y, t.m, Math.min(s.d, daysInMonth(t.y, t.m)))
    return x <= 0 || settlement >= maturity || maturity > yearOn ? E.NUM : answer(maturity - settlement, x)
  })
}

/** PRICEMAT and YIELDMAT: interest from issue paid at maturity, on year fractions from issue. */
function atMaturity(settlement: number, maturity: number, issue: number, basis: number) {
  return { dim: yearFrac(issue, maturity, basis), a: yearFrac(issue, settlement, basis), dsm: yearFrac(settlement, maturity, basis) }
}

export const functions: Record<string, FnDef> = {
  ACCRINTM: atp('e e x x? j?', (issue: number, settlement: number, rate: number, par = 1000, basis = 0) =>
    bad(issue, settlement, basis, rate, par) ? E.NUM : par * rate * yearFrac(issue, settlement, basis),
  ),
  DISC: onYears((years, pr, redemption) => (redemption - pr) / redemption / years),
  INTRATE: onYears((years, investment, redemption) => (redemption - investment) / investment / years),
  PRICEDISC: onYears((years, discount, redemption) => redemption - discount * redemption * years),
  PRICEMAT: atp('e e e x x j?', (settlement: number, maturity: number, issue: number, rate: number, yld: number, basis = 0) => {
    if (settlement >= maturity || !validBasis(basis) || rate < 0 || yld < 0) return E.NUM
    const { dim, a, dsm } = atMaturity(settlement, maturity, issue, basis)
    return ((1 + dim * rate) / (1 + dsm * yld) - a * rate) * 100
  }),
  RECEIVED: onYears((years, investment, discount) => {
    const d = 1 - discount * years
    return d <= 0 ? NaN : investment / d
  }),
  TBILLEQ: tbill((dsm, disc) => {
    const price = 100 * (1 - (disc * dsm) / 360)
    if (price <= 0) return NaN
    if (dsm <= 182) return (365 * disc) / (360 - disc * dsm)
    const t = dsm / 365
    return (-2 * t + 2 * Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (2 * t - 1)
  }),
  TBILLPRICE: tbill((dsm, disc) => {
    const price = 100 * (1 - (disc * dsm) / 360)
    return price <= 0 ? NaN : price
  }),
  TBILLYIELD: tbill((dsm, pr) => ((100 - pr) / pr) * (360 / dsm)),
  YIELDDISC: onYears((years, pr, redemption) => (redemption - pr) / pr / years),
  YIELDMAT: atp('e e e x x j?', (settlement: number, maturity: number, issue: number, rate: number, pr: number, basis = 0) => {
    if (settlement >= maturity || !validBasis(basis) || rate < 0 || pr <= 0) return E.NUM
    const { dim, a, dsm } = atMaturity(settlement, maturity, issue, basis)
    return ((1 + dim * rate) / (pr / 100 + a * rate) - 1) / dsm
  }),
}
