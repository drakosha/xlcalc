/**
 * The statistics of a rule's cells that colour scales, data bars, icon sets, top and bottom, average and duplicate
 * rules read: the numbers sorted, their mean and deviation, the count of each value. They are kept per session state
 * and rule until an edit changes a value among the rule's cells (`dropStats`).
 */
import type { ConditionalFormat } from '../contract.js'
import type { EngineState } from '../engine.js'
import { paced, type Steps } from '../exec/budget.js'
import { caseKey } from '../collate.js'
import { meets, type Rect } from '../structure/area.js'
import type { Rules } from '../locale/rules.js'
import { isErr, Rng, type V } from '../values.js'
import { rectsOf } from './rules.js'

export interface Stats {
  /** The numbers of the cells, ascending (dates are numbers; text, booleans and blanks are not). */
  nums: number[]
  mean: number
  /** The population standard deviation of the numbers. */
  sd: number
  /** How many cells hold each value (`valueKey`), blank ones too. */
  counts: Map<string, number>
}

/**
 * A value as duplicate rules compare it (the set `cond/looks`): text in any case, text that reads as a number as that
 * number (`"1"` is a duplicate of 1), empty text as a blank cell (blank cells themselves are never formatted).
 */
export const valueKey =
  (rules: Rules) =>
  (v: V): string => {
    if (v === null || v === '') return ''
    if (typeof v === 'number') return `n${v}`
    if (typeof v === 'boolean') return `b${v}`
    if (isErr(v)) return `e${v.code}`
    const n = rules.toNum(v)
    return typeof n === 'number' ? `n${n}` : `s${caseKey(v)}`
  }

/** The values of each cell of a rule's ranges, once each where ranges overlap, row by row; the steps yield between cells. */
export function* eachValue(state: EngineState, s: number, rects: Rect[], f: (v: V) => void): Steps<void> {
  for (let k = 0; k < rects.length; k++) {
    const x = rects[k]
    const before = rects.slice(0, k).filter((y) => meets(x, y))
    const g = new Rng(s, x.r1, x.c1, x.r2, x.c2)
    const b = state.rt.bounds(g)
    const flat = state.rt.flat(g)
    for (let i = 0; i < flat.length; i++) {
      if (paced(1)) yield
      if (before.length) {
        const r = b.r1 + Math.floor(i / b.w)
        const c = b.c1 + (i % b.w)
        if (before.some((y) => r >= y.r1 && r <= y.r2 && c >= y.c1 && c <= y.c2)) continue
      }
      f(flat[i])
    }
  }
}

/** The statistics of a rule on sheet `s`, from the current values. */
export function* ruleStats(state: EngineState, s: number, rule: ConditionalFormat): Steps<Stats> {
  const nums: number[] = []
  const counts = new Map<string, number>()
  const key = valueKey(state.ctx.rules)
  yield* eachValue(state, s, rectsOf(rule), (v) => {
    if (typeof v === 'number') nums.push(v)
    const k = key(v)
    counts.set(k, (counts.get(k) ?? 0) + 1)
  })
  nums.sort((a, b) => a - b)
  let sum = 0
  for (const x of nums) sum += x
  const mean = nums.length ? sum / nums.length : 0
  let sq = 0
  for (const x of nums) sq += (x - mean) ** 2
  return { nums, mean, sd: nums.length ? Math.sqrt(sq / nums.length) : 0, counts }
}

/** PERCENTILE.INC of sorted numbers, `k` from 0 to 1. */
export function percentile(nums: number[], k: number): number {
  const h = Math.min(Math.max(k, 0), 1) * (nums.length - 1)
  const i = Math.floor(h)
  return i >= nums.length - 1 ? nums[nums.length - 1] : nums[i] + (h - i) * (nums[i + 1] - nums[i])
}

/** The statistics kept per state, by sheet and rule. */
const kept = new WeakMap<EngineState, Map<ConditionalFormat, { s: number; rects: Rect[]; stats: Stats }>>()

/** A rule's statistics, kept until an edit among its cells. */
export function* statsOf(state: EngineState, s: number, rule: ConditionalFormat): Steps<Stats> {
  let m = kept.get(state)
  if (!m) kept.set(state, (m = new Map()))
  const hit = m.get(rule)
  if (hit) return hit.stats
  const stats = yield* ruleStats(state, s, rule)
  m.set(rule, { s, rects: rectsOf(rule), stats })
  return stats
}

/** Drops the statistics of rules whose cells meet the changed cells (`s:r:c` keys); all of them with `all`. */
export function dropStats(state: EngineState, keys: string[], all = false) {
  const m = kept.get(state)
  if (!m?.size) return
  if (all) return m.clear()
  for (const k of keys) {
    const [s, r, c] = k.split(':').map(Number)
    for (const [rule, x] of m) if (x.s === s && x.rects.some((y) => r >= y.r1 && r <= y.r2 && c >= y.c1 && c <= y.c2)) m.delete(rule)
  }
}
