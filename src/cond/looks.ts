/**
 * What a sheet's conditional formats give the cells of a window (`Session.getCells`), by the values the cells hold now.
 * The rules apply in the sheet's order: a field of a style an earlier rule set stays, the first colour (a style's fill
 * or a colour scale's), bar and icon win, and a rule with `stopIfTrue` that holds leaves out the rules after it.
 * Rules over the numbers of their cells (colour scales, data bars, icon sets, top and bottom, average, duplicates) read
 * the rule's statistics (`stats.ts`); rules with formulas compute them at each cell (`formulas.ts`).
 */
import type { BarLook, CellInput, CellValue, Color, ConditionalFormat, ConditionalAppearance, ConditionalStyle, ConditionalValue, DataBarFormat } from '../contract.js'
import type { EngineState } from '../engine.js'
import type { Looks } from '../shown.js'
import { paced, type Steps } from '../exec/budget.js'
import { colOf, inner, MAX_COLS, rowOf } from '../model.js'
import { enCode } from '../numfmt/builtin.js'
import { COMPARE } from '../operators.js'
import { cellCount, hasCell, inside, type Rect } from '../structure/area.js'
import type { Ast } from '../syntax.js'
import { E, isErr, type V } from '../values.js'
import { evaluator, holds, prepare } from './formulas.js'
import { interpolate } from '../functions/statistical/common.js'
import { originOf, rectsOf, ruleFormula } from './rules.js'
import { statsOf, valueKey, type Stats } from './stats.js'

/** Formulas checked once per state: by sheet and text. */
const prepared = new WeakMap<EngineState, Map<string, Ast | null>>()

function tree(state: EngineState, s: number, text: string): Ast | null {
  const m = inner(prepared, state, () => new Map())
  const k = `${s}\n${text}`
  if (!m.has(k)) m.set(k, prepare(state, s, text))
  return m.get(k)!
}

/** A formula of a rule as a function of the cell (written for the rule's first cell); an error where it does not parse. */
function formulaAt(state: EngineState, s: number, rule: ConditionalFormat, text: string): (r: number, c: number) => V {
  const ast = tree(state, s, text)
  return ast ? evaluator(state, s, ast, text, originOf(rule)) : () => E.ERROR
}

const num = (v: V): number | null => (typeof v === 'number' ? v : null)

/** A threshold's number over the rule's statistics; null when it has none (no numbers, an error among the cells it reads, a formula that is no number). */
function threshold(state: EngineState, s: number, rule: ConditionalFormat, t: ConditionalValue, st: Stats): number | null {
  const { nums } = st
  const lo = nums[0]
  const hi = nums[nums.length - 1]
  const own = () => (t.value === undefined ? null : num(formulaAt(state, s, rule, t.value)(originOf(rule).r, originOf(rule).c)))
  // errors among the cells are left out, as text and blanks are (the set `cond/looks`)
  const stat = (f: () => number) => (nums.length ? f() : null)
  switch (t.type) {
    case 'min':
      return stat(() => lo)
    case 'max':
      return stat(() => hi)
    case 'autoMin':
      return stat(() => Math.min(0, lo))
    case 'autoMax':
      return stat(() => Math.max(0, hi))
    case 'num':
    case 'formula':
      return own()
    case 'percent': {
      const p = own()
      return p === null ? null : stat(() => lo + ((hi - lo) * p) / 100)
    }
    case 'percentile': {
      const p = own()
      return p === null || p < 0 || p > 100 ? null : stat(() => interpolate(nums, (p / 100) * (nums.length - 1)))
    }
  }
}

const channels = (c: Color) => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16))
const hex = (xs: number[]) => `#${xs.map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`

/** The colour a fraction `t` of the way from `a` to `b`: each channel moves by its share cut toward `a` (the set `cond/looks`). */
export const between = (a: Color, b: Color, t: number): Color => {
  const [x, y] = [channels(a), channels(b)]
  return hex(x.map((v, i) => v + Math.trunc((y[i] - v) * t)))
}

/** A colour scale's colour for `v` over its points' numbers `at` (ascending as the rule gives them). */
function scaleColor(v: number, at: number[], colors: Color[]): Color {
  const n = at.length
  if (v >= at[n - 1]) return colors[n - 1]
  if (v <= at[0]) return colors[0]
  for (let i = 1; i < n; i++) if (v <= at[i]) return at[i] === at[i - 1] ? colors[i] : between(colors[i - 1], colors[i], (v - at[i - 1]) / (at[i] - at[i - 1]))
  return colors[n - 1]
}

const clamp = (x: number) => Math.min(Math.max(x, 0), 1)

/**
 * A data bar for `v` between `lo` and `hi`: Excel 2010's from the axis (at the numbers' zero when they reach below it,
 * the middle with `axis: 'middle'`), lengths from `minLength` to `maxLength` percent; Excel 2007's from the left edge.
 */
export function barOf(rule: DataBarFormat, v: number, lo: number, hi: number): BarLook {
  const old = rule.excel2007 === true
  const minL = (rule.minLength ?? (old ? 10 : 0)) / 100
  const maxL = (rule.maxLength ?? (old ? 90 : 100)) / 100
  const span = (f: number) => minL + clamp(f) * (maxL - minL)
  const look = (start: number, end: number, negative = false): BarLook => {
    const out: BarLook = { start, end, color: negative && !old ? (rule.negativeColor ?? rule.color) : rule.color, gradient: old || rule.gradient !== false }
    const border = old ? undefined : negative ? (rule.negativeBorder ?? rule.border) : rule.border
    if (border) out.border = border
    if (!old && rule.direction === 'rightToLeft') out.rightToLeft = true
    return out
  }
  const f = hi === lo ? 1 : (v - lo) / (hi - lo)
  if (old || rule.axis === 'none' || (lo >= 0 && rule.axis !== 'middle')) return look(0, span(f), v < 0)
  // an axis: at the numbers' zero, or the middle; each side as long as its end is far from it
  const axis = rule.axis === 'middle' ? 0.5 : hi <= 0 ? 1 : clamp(-lo / (hi - lo))
  const out = v >= 0 ? look(axis, axis + (1 - axis) * span(hi > 0 ? v / (rule.axis === 'middle' ? Math.max(hi, -lo) : hi) : 0)) : look(axis - axis * span(lo < 0 ? v / (rule.axis === 'middle' ? Math.min(lo, -hi) : lo) : 0), axis, true)
  out.axis = axis
  out.axisColor = rule.axisColor ?? '#000000'
  return out
}

/** A rule's style fields merged under the fields set already (an earlier rule's win). */
function mergeStyle(into: ConditionalAppearance, s: ConditionalStyle) {
  const out = (into.style ??= {})
  if (s.font) {
    const font: Record<string, unknown> = (out.font ??= {})
    for (const [k, v] of Object.entries(s.font)) if (v !== undefined && font[k] === undefined) font[k] = v
  }
  if (s.fill && !out.fill) out.fill = s.fill
  if (s.border) {
    const border: Record<string, unknown> = (out.border ??= {})
    for (const [k, v] of Object.entries(s.border)) if (v && !border[k]) border[k] = v
  }
  if (s.numberFormat !== undefined && out.numberFormat === undefined) out.numberFormat = s.numberFormat
}

/** Whether a rule of a style holds for `v` at (r, c). */
type Test = (v: V, r: number, c: number) => boolean

/** The kinds that shade a cell by where its number sits among the range's numbers instead of testing it for a style: `shade` takes them, `styleTest` the rest. */
type Graded = { type: 'colorScale' | 'dataBar' | 'iconSet' }

/** The test of a rule with a style, built once per window; null when it holds for no cell. */
function* styleTest(state: EngineState, s: number, rule: Exclude<ConditionalFormat, Graded>): Steps<Test | null> {
  const { compareOp } = state.ctx.rules
  const at = (text: string) => formulaAt(state, s, rule, text)
  switch (rule.type) {
    case 'cellIs': {
      const [fa, fb] = rule.formulas.map(at)
      // a formula's answer as a cell shows it: a blank cell it refers to is 0 (the set `cond/looks`)
      const cmp = (v: V, x: V) => (isErr(x) ? NaN : compareOp(v, x ?? 0))
      const op = COMPARE[rule.operator]
      return (v, r, c) => {
        if (isErr(v)) return false
        const a = cmp(v, fa(r, c))
        if (Number.isNaN(a)) return false
        if (rule.operator !== 'between' && rule.operator !== 'notBetween') return op(a, 0)
        const b = cmp(v, fb(r, c))
        if (Number.isNaN(b)) return false
        // the ends either way round (the set `cond/looks`, column E), where data validation's take nothing reversed
        const inside = COMPARE.between(a, b) || COMPARE.between(b, a)
        return rule.operator === 'between' ? inside : !inside
      }
    }
    case 'containsBlanks':
    case 'notContainsBlanks': {
      // by the value, not the formula Excel writes (`LEN(TRIM(A1))>0`): an error is not blank (the set `cond/looks`)
      const blank = rule.type === 'containsBlanks'
      return (v) => (v === null || (typeof v === 'string' && /^ *$/.test(v))) === blank
    }
    case 'top10': {
      const st = yield* statsOf(state, s, rule)
      const n = st.nums.length
      if (!n) return null
      const k = Math.min(n, rule.percent ? Math.max(1, Math.floor((n * rule.rank) / 100)) : rule.rank)
      const edge = rule.bottom ? st.nums[k - 1] : st.nums[n - k]
      return (v) => typeof v === 'number' && (rule.bottom ? v <= edge : v >= edge)
    }
    case 'aboveAverage': {
      const st = yield* statsOf(state, s, rule)
      if (!st.nums.length) return null
      const d = (rule.stdDev ?? 0) * st.sd
      const below = rule.below === true
      const edge = below ? st.mean - d : st.mean + d
      return (v) => typeof v === 'number' && (below ? (rule.equal ? v <= edge : v < edge) : rule.equal ? v >= edge : v > edge)
    }
    case 'duplicateValues':
    case 'uniqueValues': {
      const st = yield* statsOf(state, s, rule)
      const dup = rule.type === 'duplicateValues'
      const key = valueKey(state.ctx.rules)
      return (v) => v !== null && (st.counts.get(key(v)) ?? 0) > 1 === dup
    }
    default: {
      const f = formulaAt(state, s, rule, ruleFormula(rule)!)
      return (_, r, c) => holds(state, f(r, c))
    }
  }
}

/** What a rule over numbers gives a number, built once per window; null when it gives nothing. */
type Shade = (v: number, into: ConditionalAppearance) => void

function* shade(state: EngineState, s: number, rule: Extract<ConditionalFormat, Graded>): Steps<Shade | null> {
  const st = yield* statsOf(state, s, rule)
  const t = (v: ConditionalValue) => threshold(state, s, rule, v, st)
  switch (rule.type) {
    case 'colorScale': {
      const at = rule.points.map((p) => t(p.value))
      if (at.some((x) => x === null)) return null
      const colors = rule.points.map((p) => p.color)
      return (v, into) => {
        if (into.style?.fill) return
        ;(into.style ??= {}).fill = { color: scaleColor(v, at as number[], colors) }
      }
    }
    case 'dataBar': {
      const [lo, hi] = [t(rule.min), t(rule.max)]
      if (lo === null || hi === null) return null
      return (v, into) => {
        if (into.bar) return
        into.bar = barOf(rule, v, lo, hi)
        if (rule.showValue === false) into.hideValue = true
      }
    }
    case 'iconSet': {
      const at = rule.points.map((p) => t(p.value))
      if (at.slice(1).some((x) => x === null)) return null
      const n = rule.points.length
      return (v, into) => {
        if (into.icon !== undefined) return
        let k = 0
        for (let i = n - 1; i >= 1; i--)
          if (rule.points[i].gte === false ? v > at[i]! : v >= at[i]!) {
            k = i
            break
          }
        const own = rule.points[k].icon
        if (own === null) return
        into.icon = own ?? { set: rule.icons, index: rule.reverse ? n - 1 - k : k }
        if (rule.showValue === false) into.hideValue = true
      }
    }
  }
}

/**
 * The looks of the cells of rectangle `x` of sheet `s` (of each of several, apart), by `r * MAX_COLS + c`; `only`, the
 * cells to look at when the window is too large to look at every one (the cells it lists). The steps yield between cells.
 */
export function* windowLooks(state: EngineState, s: number, x: Rect | Rect[], only?: Set<number>): Steps<Map<number, ConditionalAppearance>> {
  const out = new Map<number, ConditionalAppearance>()
  const rules = state.input.sheets[s].conditionalFormats
  if (!rules?.length) return out
  const stopped = new Set<number>()
  const value = (r: number, c: number) => state.rt.cell(s, r, c)
  const xs = Array.isArray(x) ? x : [x]
  for (const rule of rules) {
    // the part of a rule's ranges inside the window's rectangles
    const parts = rectsOf(rule).flatMap((y) => inside(xs, y))
    if (!parts.length) continue
    const graded = rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet'
    const test = graded ? null : yield* styleTest(state, s, rule)
    const paint = graded ? yield* shade(state, s, rule) : null
    if (!test && !paint) continue
    const style = 'style' in rule ? rule.style : undefined
    const visit = (k: number, r: number, c: number) => {
      if (stopped.has(k)) return
      const v = value(r, c)
      if (paint) {
        if (typeof v !== 'number') return
        const look = out.get(k) ?? {}
        paint(v, look)
        if (Object.keys(look).length) out.set(k, look)
        return
      }
      if (!test!(v, r, c)) return
      if (rule.stopIfTrue) stopped.add(k)
      if (!style) return
      const look = out.get(k) ?? {}
      mergeStyle(look, style)
      out.set(k, look)
    }
    // the cells listed, each tested against the parts, when that is less work than the parts' every cell
    if (only && only.size * parts.length < parts.reduce((n, p) => n + cellCount(p), 0)) {
      for (const k of only) {
        if (paced(1)) yield
        const r = rowOf(k)
        const c = colOf(k)
        if (parts.some((p) => hasCell(p, r, c))) visit(k, r, c)
      }
      continue
    }
    const seen = parts.length > 1 ? new Set<number>() : null
    for (const p of parts)
      for (let r = p.r1; r <= p.r2; r++)
        for (let c = p.c1; c <= p.c2; c++) {
          if (paced(1)) yield
          const k = r * MAX_COLS + c
          if ((only && !only.has(k)) || seen?.has(k)) continue
          seen?.add(k)
          visit(k, r, c)
        }
  }
  return out
}

/** What shows a value by a format: the view's text and colour by it. */
type ShownBy = ((format: NonNullable<ConditionalStyle['numberFormat']>) => { text: string; color?: Color }) | null | undefined

/**
 * A view's text and colour by its look's number format, when it has one (`display`, the view's text by a format). A
 * built-in format given by number shows by its en-US code in every locale, as Excel reads a differential format by the
 * code it holds (the set `cond/formats-ru-held`: 14 shows `m/d/yyyy` in ru-RU too, `3.15.2026`).
 */
export function lookShown(view: { text?: string; color?: Color }, look: ConditionalAppearance, display: ShownBy) {
  const format = look.style?.numberFormat
  if (format === undefined || !display) return
  const d = display(enCode(format))
  view.text = d.text
  if (d.color === undefined) delete view.color
  else view.color = d.color
}

/** A cell's view with its conditional formats' appearance; a number format of theirs shows its value by it. */
export function withLook(view: { text?: string; color?: Color; appearance?: ConditionalAppearance }, look: ConditionalAppearance, display?: ShownBy) {
  view.appearance = look
  lookShown(view, look, display)
}

/**
 * What a window puts on a cell of sheet `s` (`WindowSource.appearance`): the appearance `cond` holds for it
 * (`windowLooks`), its value shown by a number format of theirs through `looks` (null without the display).
 */
export const appearanceOf =
  (cond: Map<number, ConditionalAppearance>, looks: Looks | null, s: number) =>
  (view: { value: CellValue; text?: string; color?: Color; appearance?: ConditionalAppearance }, r: number, c: number, own: CellInput | undefined) => {
    const look = cond.get(r * MAX_COLS + c)
    if (look) withLook(view, look, looks && ((format) => looks(s, r, c, view.value, { ...own, numberFormat: format })))
  }
