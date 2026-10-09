/** Scalar aggregates over shared range prefixes, with unrounded sums and the original row-major order. */
import type { Bounds } from './args.js'
import { criterionMatcher, errorCrit, matches, parseCrit, type CritRules } from './criteria.js'
import { CompensatedSum, meanOf } from './functions/statistical/sum.js'
import { readsClock } from './locale/rules/data.js'
import type { CacheBudget } from './range-cache.js'
import { RangePrefixes } from './range-prefix.js'
import type { Area } from './spill/state.js'
import { E, fitNumber, isErr, type Rng, type CErr, type V } from './values.js'

export type AggregateName = 'SUM' | 'COUNT' | 'COUNTA' | 'MIN' | 'MAX' | 'AVERAGE'
interface Partial {
  sum: number
  count: number
  filled: number
  min: number
  max: number
  error?: CErr
}
interface Mean {
  sum: CompensatedSum
  count: number
  first: number
  constant: boolean
  error?: CErr
}
interface Count { count: number; seen: number }
type Cell = (s: number, r: number, c: number) => V
const CRITERIA = 64

function addNumber(result: Partial, value: V): void {
  if (value !== null) result.filled++
  if (typeof value === 'number') {
    result.sum += value
    result.count++
    if (value < result.min) result.min = value
    if (value > result.max) result.max = value
  } else if (isErr(value)) result.error ??= value
}

function addMean(result: Mean, value: V): void {
  if (typeof value === 'number') {
    result.sum.add(value)
    if (result.count++ === 0) result.first = value
    result.constant &&= value === result.first
  } else if (isErr(value)) result.error ??= value
}

export class RangeAggregates {
  private readonly numbers: RangePrefixes<Partial>
  private readonly means: RangePrefixes<Mean>
  private readonly criteria = new Map<string, RangePrefixes<Count>>()

  constructor(private readonly budget: CacheBudget) {
    this.numbers = new RangePrefixes<Partial>({
      create: () => ({ sum: 0, count: 0, filled: 0, min: Infinity, max: -Infinity }),
      copy: (value: Partial) => ({ ...value }), bytes: 96,
    }, budget)
    this.means = new RangePrefixes<Mean>({
      create: () => ({ sum: new CompensatedSum(), count: 0, first: 0, constant: true }),
      copy: (value: Mean) => ({ ...value, sum: value.sum.copy() }), bytes: 128,
    }, budget)
  }

  invalidate(area: Area): void {
    this.numbers.invalidate(area)
    this.means.invalidate(area)
    for (const prefix of this.criteria.values()) prefix.invalidate(area)
  }

  /** Drop scalar summaries before a large batch writes its formulas, avoiding one full prefix scan per cell. */
  invalidateCells(cells: readonly { s: number; r: number; c: number }[]): void {
    const areas = new Map<number, Area>()
    for (const { s, r, c } of cells) {
      const a = areas.get(s)
      if (!a) areas.set(s, { s, r1: r, r2: r, c1: c, c2: c })
      else {
        a.r1 = Math.min(a.r1, r); a.r2 = Math.max(a.r2, r)
        a.c1 = Math.min(a.c1, c); a.c2 = Math.max(a.c2, c)
      }
    }
    for (const a of areas.values()) this.invalidate(a)
  }

  /** Undefined retains AVERAGE's existing exact fallback when the compensated sum overflows. */
  read(name: AggregateName, range: Rng, bounds: () => Bounds, cell: Cell): V | undefined {
    if (name === 'AVERAGE') {
      const result = this.means.read(range, bounds, cell, addMean)
      if (result.error) return result.error
      if (!result.count) return E.DIV0
      const value = meanOf(result.sum, result.count, result.first, result.constant)
      return value === undefined ? undefined : fitNumber(value)
    }
    const result = this.numbers.read(range, bounds, cell, addNumber)
    if (name === 'COUNT') return result.count
    if (name === 'COUNTA') return result.filled
    if (result.error) return result.error
    return fitNumber(name === 'SUM' ? result.sum : result.count === 0 ? 0 : name === 'MIN' ? result.min : result.max)
  }

  countIf(range: Rng, value: V, rules: CritRules, bounds: () => Bounds, cell: Cell): number {
    const key = isErr(value) ? `error:${value.code}` : `${typeof value}:${String(value)}`
    let prefix = this.criteria.get(key)
    if (prefix) this.criteria.delete(key)
    else {
      if (this.criteria.size === CRITERIA) {
        const first = this.criteria.keys().next().value!
        this.criteria.get(first)!.clear()
        this.criteria.delete(first)
      }
      prefix = new RangePrefixes<Count>({ create: () => ({ count: 0, seen: 0 }), copy: (p: Count) => ({ ...p }), bytes: 64 }, this.budget)
    }
    this.criteria.set(key, prefix)
    const [answer, clocked] = readsClock(rules.clock, () => {
      const crit = isErr(value) ? errorCrit(value.code) : parseCrit(value, rules)
      const test = criterionMatcher(crit, rules)
      const result = prefix.read(range, bounds, cell, (p, v) => { p.seen++; if (test(v)) p.count++ })
      const size = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1)
      return result.count + (matches(null, crit, rules) ? size - result.seen : 0)
    })
    // Both criterion text and matched cell text can read the current year; every caller must keep that dependency.
    if (clocked) { prefix.clear(); this.criteria.delete(key) }
    return answer
  }
}
