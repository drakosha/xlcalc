/**
 * Conditional formats under structure operations. A rule's ranges change as merged areas do: they grow and shrink with
 * inserted and deleted lines, go with moved cells and are copied with copied ones, and the cells a move or a copy
 * lands on lose the rules they had. Its formulas change as the formula of a cell at the rule's first cell would, and
 * when that cell is gone or no longer first they are made for the new first cell as a copy would make them.
 */
import type { BookInput, ConditionalFormat, ConditionalValue, SheetInput } from '../contract.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { sheetKey } from '../syntax.js'
import { meets, type Axis, type Rect } from '../structure/area.js'
import { copyTransform } from '../structure/copy.js'
import type { Paste } from '../structure/paste.js'
import { rewriteFormula, type Transform } from '../structure/rewrite.js'
import { deleteSpan, insertSpan } from '../structure/shift.js'
import { rangesOf, rectsOf } from './rules.js'

type Cell = { r: number; c: number }

const valueWith = (v: ConditionalValue, f: (text: string) => string): ConditionalValue => (v.value === undefined ? v : { ...v, value: f(v.value) })

/** The rule with each of its formulas through `f`. */
function mapFormulas(rule: ConditionalFormat, f: (text: string) => string): ConditionalFormat {
  switch (rule.type) {
    case 'cellIs':
      return { ...rule, formulas: rule.formulas.map(f) }
    case 'colorScale':
      return { ...rule, points: rule.points.map((p) => ({ ...p, value: valueWith(p.value, f) })) }
    case 'dataBar':
      return { ...rule, min: valueWith(rule.min, f), max: valueWith(rule.max, f) }
    case 'iconSet':
      return { ...rule, points: rule.points.map((p) => ({ ...p, value: valueWith(p.value, f) })) }
    default:
      return 'formula' in rule && rule.formula !== undefined ? { ...rule, formula: f(rule.formula) } : rule
  }
}

/** A rule over `next` (rectangles of the sheet after the operation), its formulas made for the new first cell: copied there from the old one by `from` (that cell's place before the operation), then rewritten by `t`. */
function placed(rule: ConditionalFormat, next: Rect[], from: Cell, host: string, moved: string, t: Transform, names: Map<string, string>): ConditionalFormat {
  const o = rectsOf(rule)[0]
  const [dr, dc] = [from.r - o.r1, from.c - o.c1]
  const copy = dr || dc ? copyTransform(dr, dc) : null
  const out = mapFormulas(rule, (text) => rewriteFormula(copy ? rewriteFormula(text, host, host, copy, names) : text, host, moved, t, names))
  const ranges = rangesOf(next, rule.ranges)
  return ranges === rule.ranges && sameFormulas(out, rule) ? rule : { ...out, ranges }
}

const sameFormulas = (a: ConditionalFormat, b: ConditionalFormat) => JSON.stringify(a) === JSON.stringify(b)

/** A sheet's rules with their formulas rewritten by `t` (cells moved elsewhere, sheets renamed or deleted); the same array when nothing changed. */
export function rewriteConditional(list: ConditionalFormat[] | undefined, host: string, moved: string, t: Transform, names: Map<string, string>): ConditionalFormat[] | undefined {
  if (!list) return list
  let changed = false
  const out = list.map((rule) => {
    const next = mapFormulas(rule, (text) => rewriteFormula(text, host, moved, t, names))
    if (sameFormulas(next, rule)) return rule
    changed = true
    return next
  })
  return changed ? out : list
}

/** The sheet with `conditionalFormats` set; left out when empty. */
function withRules(s: SheetInput, list: ConditionalFormat[]): SheetInput {
  const { conditionalFormats: _, ...rest } = s
  return list.length ? { ...rest, conditionalFormats: list } : rest
}

/**
 * The rules of the book's sheets after `n` lines are inserted (`n` > 0) or deleted at `at` of `target` (`t`, the
 * references' transform): the target's ranges shift as ranges do and a rule with none left is gone; every rule's
 * formulas are rewritten.
 */
export function shiftConditional(sheets: SheetInput[], target: string, axis: Axis, at: number, n: number, t: Transform, names: Map<string, string>): SheetInput[] {
  const max = axis === 'r' ? MAX_ROWS : MAX_COLS
  const count = Math.abs(n)
  const span = n > 0 ? insertSpan : deleteSpan
  // a line of the sheet after the operation where it was before
  const back = (i: number) => (n > 0 ? (i >= at + count ? i - count : i) : i >= at ? i + count : i)
  return sheets.map((s) => {
    const list = s.conditionalFormats
    if (!list?.length) return s
    if (sheetKey(s.name) !== sheetKey(target)) {
      const next = rewriteConditional(list, s.name, s.name, t, names)
      return next === list ? s : withRules(s, next!)
    }
    let changed = false
    const out: ConditionalFormat[] = []
    for (const rule of list) {
      const next: Rect[] = []
      for (const x of rectsOf(rule)) {
        const lohi = span(axis === 'r' ? x.r1 : x.c1, axis === 'r' ? x.r2 : x.c2, at, count, max, false)
        if (lohi) next.push(axis === 'r' ? { ...x, r1: lohi[0], r2: lohi[1] } : { ...x, c1: lohi[0], c2: lohi[1] })
      }
      if (!next.length) {
        changed = true
        continue
      }
      const from = axis === 'r' ? { r: back(next[0].r1), c: next[0].c1 } : { r: next[0].r1, c: back(next[0].c1) }
      const r = placed(rule, next, from, s.name, s.name, t, names)
      changed ||= r !== rule
      out.push(r)
    }
    return changed ? withRules(s, out) : s
  })
}

/** Rectangles without the cells of `cut`. */
export function subtract(rects: Rect[], cut: Rect): Rect[] {
  const out: Rect[] = []
  for (const x of rects) {
    if (!meets(x, cut)) {
      out.push(x)
      continue
    }
    if (x.r1 < cut.r1) out.push({ ...x, r2: cut.r1 - 1 })
    if (x.r2 > cut.r2) out.push({ ...x, r1: cut.r2 + 1 })
    const [r1, r2] = [Math.max(x.r1, cut.r1), Math.min(x.r2, cut.r2)]
    if (x.c1 < cut.c1) out.push({ r1, r2, c1: x.c1, c2: cut.c1 - 1 })
    if (x.c2 > cut.c2) out.push({ r1, r2, c1: cut.c2 + 1, c2: x.c2 })
  }
  return out
}

const inside = (x: Rect, y: Rect): Rect | null => {
  const z = { r1: Math.max(x.r1, y.r1), r2: Math.min(x.r2, y.r2), c1: Math.max(x.c1, y.c1), c2: Math.min(x.c2, y.c2) }
  return z.r1 <= z.r2 && z.c1 <= z.c2 ? z : null
}
const shifted = (x: Rect, dr: number, dc: number): Rect => ({ r1: x.r1 + dr, r2: x.r2 + dr, c1: x.c1 + dc, c2: x.c2 + dc })
const sheetOf = (book: BookInput, name: string) => book.sheets.find((s) => sheetKey(s.name) === sheetKey(name))!

/**
 * The rules after a move (`p`, cut and paste; `t`, the references' transform), set on the sheets of `out`: the moved
 * cells take their rules along (to another sheet as rules of their own there), the cells landed on lose theirs, and
 * every rule's formulas are rewritten.
 */
export function moveConditional(old: BookInput, out: BookInput, p: Paste, t: Transform, names: Map<string, string>) {
  const [fk, tk] = [sheetKey(p.from), sheetKey(p.to)]
  const [dr, dc] = p.offsets[0]
  const dst = p.area
  const carried: ConditionalFormat[] = []
  out.sheets = out.sheets.map((s) => {
    const k = sheetKey(s.name)
    const list = sheetOf(old, s.name)?.conditionalFormats
    if (k !== fk && k !== tk) {
      const next = rewriteConditional(list, s.name, s.name, t, names)
      return next === list ? s : withRules(s, next ?? [])
    }
    const rules: ConditionalFormat[] = []
    for (const rule of list ?? []) {
      const rects = rectsOf(rule)
      const stay = subtract(k === fk ? rects.flatMap((x) => subtract([x], p.src)) : rects, k === tk ? dst : { r1: -1, r2: -1, c1: -1, c2: -1 })
      const moved = k === fk ? rects.flatMap((x) => inside(x, p.src) ?? []).map((x) => shifted(x, dr, dc)) : []
      const back = (c: Cell) => (moved.some((x) => c.r >= x.r1 && c.r <= x.r2 && c.c >= x.c1 && c.c <= x.c2) ? { r: c.r - dr, c: c.c - dc } : c)
      if (fk === tk) {
        const next = [...stay, ...moved]
        if (next.length) rules.push(placed(rule, next, back({ r: next[0].r1, c: next[0].c1 }), s.name, s.name, t, names))
        continue
      }
      if (stay.length) rules.push(placed(rule, stay, { r: stay[0].r1, c: stay[0].c1 }, s.name, s.name, t, names))
      if (moved.length) carried.push(placed(rule, moved, { r: moved[0].r1 - dr, c: moved[0].c1 - dc }, p.from, p.to, t, names))
    }
    return withRules(s, rules)
  })
  if (carried.length) out.sheets = out.sheets.map((s) => (sheetKey(s.name) === tk ? withRules(s, [...(s.conditionalFormats ?? []), ...carried]) : s))
}

/**
 * The rules after a copy (`p`, copy and paste or fill; `source`, the book copied from): the cells pasted on lose their
 * rules and take those of the cells copied, a rule of the same sheet growing over the copies, one of another sheet
 * becoming a rule of the target's, first in its order, its formulas copied as the cells' are.
 */
export function copyConditional(source: BookInput, out: BookInput, p: Paste, names: Map<string, string>) {
  const [fk, tk] = [sheetKey(p.from), sheetKey(p.to)]
  const target = sheetOf(out, p.to)
  const from = sheetOf(source, p.from).conditionalFormats ?? []
  const pasted = (rule: ConditionalFormat) => p.offsets.flatMap(([dr, dc]) => rectsOf(rule).flatMap((x) => inside(x, p.src) ?? []).map((x) => shifted(x, dr, dc)))
  const own = target.conditionalFormats ?? []
  if (!own.length && !from.some((r) => rectsOf(r).some((x) => meets(x, p.src)))) return
  const kept: ConditionalFormat[] = []
  for (const rule of own) {
    const stay = subtract(rectsOf(rule), p.area)
    const grown = fk === tk ? [...stay, ...pasted(rule)] : stay
    if (!grown.length) continue
    const ranges = rangesOf(grown, rule.ranges)
    kept.push(ranges === rule.ranges ? rule : { ...rule, ranges })
  }
  const added: ConditionalFormat[] = []
  if (fk !== tk)
    for (const rule of from) {
      const next = pasted(rule)
      if (!next.length) continue
      const o = rectsOf(rule)[0]
      const copy = copyTransform(next[0].r1 - o.r1, next[0].c1 - o.c1)
      added.push({ ...mapFormulas(rule, (text) => rewriteFormula(text, p.to, p.to, copy, names)), ranges: rangesOf(next, []) })
    }
  const list = [...added, ...kept]
  out.sheets = out.sheets.map((s) => (s === target ? withRules(s, list) : s))
}
