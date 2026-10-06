/** Securities without periodic coupons: discounted, interest at maturity, Treasury bills. */
import type { FnDef } from '../../args.js'
import { serialOf, ymd } from '../../dates.js'
import { E } from '../../values.js'
import { atp } from '../date/atp.js'
import { basisFirst, validBasis, yearFrac, yearLength } from '../date/daycount.js'

/** settlement < maturity, a valid basis and positive amounts, as the ToolPak checks them. */
const bad = (settlement: number, maturity: number, basis: number, ...positive: number[]) =>
  settlement >= maturity || !validBasis(basis) || positive.some((x) => x <= 0)

/** Securities of (settlement, maturity, a, b, basis) on the year fraction between the dates. */
function onYears(answer: (years: number, a: number, b: number) => number) {
  return atp('e e x x j?', (settlement: number, maturity: number, a: number, b: number, basis = 0) =>
    bad(settlement, maturity, basis, a, b) ? E.NUM : answer(yearFrac(settlement, maturity, basis), a, b),
  basisFirst(4))
}

/** The same day a year on, an impossible day rolling over (2008-02-29 gives 2009-03-01). */
const yearOn = (serial: number) => {
  const s = ymd(serial)
  return serialOf(s.y + 1, s.m, s.d)
}

/** Treasury bills: settlement < maturity ≤ a year on, the third argument positive; days are actual. */
function tbill(answer: (dsm: number, x: number) => number) {
  return atp('e e x', (settlement: number, maturity: number, x: number) =>
    x <= 0 || settlement >= maturity || maturity > yearOn(settlement) ? E.NUM : answer(maturity - settlement, x),
  )
}

/** PRICEMAT and YIELDMAT: interest from issue paid at maturity; on basis 1 every span is divided by the year of issue to settlement. */
function atMaturity(settlement: number, maturity: number, issue: number, basis: number) {
  if (basis !== 1) return { dim: yearFrac(issue, maturity, basis), a: yearFrac(issue, settlement, basis), dsm: yearFrac(settlement, maturity, basis) }
  const b = yearLength(issue, settlement)
  return { dim: (maturity - issue) / b, a: (settlement - issue) / b, dsm: (maturity - settlement) / b }
}

export const functions: Record<string, FnDef> = {
  // par is required: Excel rejects ACCRINTM with three arguments
  ACCRINTM: atp('e e x x j?', (issue: number, settlement: number, rate: number, par: number, basis = 0) =>
    bad(issue, settlement, basis, rate, par) ? E.NUM : par * rate * yearFrac(issue, settlement, basis),
  ),
  DISC: onYears((years, pr, redemption) => (redemption - pr) / redemption / years),
  INTRATE: onYears((years, investment, redemption) => (redemption - investment) / investment / years),
  PRICEDISC: onYears((years, discount, redemption) => redemption - discount * redemption * years),
  PRICEMAT: atp('e e e x x j?', (settlement: number, maturity: number, issue: number, rate: number, yld: number, basis = 0) => {
    if (settlement >= maturity || issue > settlement || !validBasis(basis) || rate < 0 || yld < 0) return E.NUM
    const { dim, a, dsm } = atMaturity(settlement, maturity, issue, basis)
    return ((1 + dim * rate) / (1 + dsm * yld) - a * rate) * 100
  }, basisFirst(5)),
  RECEIVED: onYears((years, investment, discount) => {
    const d = 1 - discount * years
    return d <= 0 ? NaN : investment / d
  }),
  // over half a year the bond-equivalent yield, with at most one year of 365 days
  TBILLEQ: tbill((dsm, disc) => {
    const price = 100 * (1 - (disc * dsm) / 360)
    if (price <= 0) return NaN
    if (dsm <= 182) return (365 * disc) / (360 - disc * dsm)
    const t = Math.min(dsm / 365, 1)
    return (-2 * t + 2 * Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (2 * t - 1)
  }),
  TBILLPRICE: tbill((dsm, disc) => {
    const price = 100 * (1 - (disc * dsm) / 360)
    return price <= 0 ? NaN : price
  }),
  TBILLYIELD: tbill((dsm, pr) => ((100 - pr) / pr) * (360 / dsm)),
  YIELDDISC: onYears((years, pr, redemption) => (redemption - pr) / pr / years),
  YIELDMAT: atp('e e e x x j?', (settlement: number, maturity: number, issue: number, rate: number, pr: number, basis = 0) => {
    if (settlement >= maturity || issue > settlement || !validBasis(basis) || rate < 0 || pr <= 0) return E.NUM
    const { dim, a, dsm } = atMaturity(settlement, maturity, issue, basis)
    return ((1 + dim * rate) / (pr / 100 + a * rate) - 1) / dsm
  }, basisFirst(5)),
}
