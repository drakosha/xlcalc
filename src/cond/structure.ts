/**
 * Conditional formats under structure operations, as Excel changes them (the set `cond/struct-held`). A rule's ranges grow
 * and shrink with inserted and deleted lines, go with moved cells and are copied with copied ones. Its formulas change as
 * the formula of a cell at the rule's first cell (the top-left of the rectangle around its ranges) would, and when that
 * cell is no longer first they are made for the new first cell as a copy would make them. Excel also changes the order
 * of rules: some it takes out and puts back after the others, and pasted rules come first.
 */
import type { BookInput, ConditionalFormat, ConditionalValue, SheetInput } from '../contract.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { sheetKey } from '../syntax.js'
import { meets, type Axis, type Rect } from '../structure/area.js'
import { copyTransform } from '../structure/copy.js'
import type { Paste } from '../structure/paste.js'
import { rewriteFormula, type Transform } from '../structure/rewrite.js'
import { deleteSpan } from '../structure/shift.js'
import { tokenize } from '../syntax.js'
import { originOf, rangesOf, rectsOf } from './rules.js'

type Cell = { r: number; c: number }

/** The top-left of the rectangle around rectangles: the cell a rule over them has its formulas written for (`originOf`). */
const corner = (xs: Rect[]): Cell => ({ r: Math.min(...xs.map((x) => x.r1)), c: Math.min(...xs.map((x) => x.c1)) })

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

/** A rule placed by an operation, and whether the operation's transform rewrote its formulas. */
interface Placed {
  rule: ConditionalFormat
  rewrote: boolean
}

const texts = (r: ConditionalFormat): string => JSON.stringify(r)

/**
 * A rule over `next` (rectangles of the sheet after the operation), its formulas made for its new first cell (`corner`):
 * copied from the old one to `from`, that cell's place before the operation, then rewritten by `t`.
 */
function placed(rule: ConditionalFormat, next: Rect[], from: Cell, host: string, moved: string, t: Transform, names: Map<string, string>): Placed {
  const o = originOf(rule)
  const [dr, dc] = [from.r - o.r, from.c - o.c]
  const copy = dr || dc ? copyTransform(dr, dc) : null
  const anchored = copy ? mapFormulas(rule, (text) => rewriteFormula(text, host, host, copy, names)) : rule
  const out = mapFormulas(anchored, (text) => plainLost(rewriteFormula(text, host, moved, t, names)))
  const ranges = rangesOf(next, rule.ranges)
  return { rule: ranges === rule.ranges && texts(out) === texts(rule) ? rule : { ...out, ranges }, rewrote: texts(out) !== texts(anchored) }
}

/** A reference to a deleted sheet in a rule's formula, `#REF!A1`, as Excel writes it there: `#REF!` (the set `cond/struct-held`). */
function plainLost(text: string): string {
  if (!text.includes('#REF!')) return text
  let toks: ReturnType<typeof tokenize>
  try {
    toks = tokenize(text)
  } catch {
    return text
  }
  let out = text
  for (const tok of toks.reverse()) if (tok.t === 'err' && tok.ref && tok.ref !== '#' && tok.sheet === undefined) out = out.slice(0, tok.at + 5) + out.slice(tok.at + 5 + tok.ref.length)
  return out
}

/** A sheet's rules with their formulas rewritten by `t` (cells moved elsewhere, sheets renamed or deleted); the same array when nothing changed. */
export function rewriteConditional(list: ConditionalFormat[] | undefined, host: string, moved: string, t: Transform, names: Map<string, string>): ConditionalFormat[] | undefined {
  if (!list) return list
  let changed = false
  const out = list.map((rule) => {
    const next = mapFormulas(rule, (text) => plainLost(rewriteFormula(text, host, moved, t, names)))
    if (texts(next) === texts(rule)) return rule
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

/** Rules in their order, those Excel takes out and puts back after the others last (`late`), in their order. */
const lastLate = (xs: { rule: ConditionalFormat; late: boolean }[]) => [...xs.filter((x) => !x.late), ...xs.filter((x) => x.late)].map((x) => x.rule)

const GRADED = new Set(['colorScale', 'dataBar', 'iconSet'])

/**
 * One axis of a rule's range after `n` lines are inserted before `at`: the lines take the format of the line before them,
 * so a range holding that line grows over them (the set `cond/struct-held`: rows inserted right below a range join it);
 * a range at or after `at` moves; a whole column or row stays. Null when it leaves the sheet.
 */
function grown(lo: number, hi: number, at: number, n: number, max: number): [number, number] | null {
  if (lo === 0 && hi === max - 1) return [lo, hi]
  if (lo >= at) return lo + n > max - 1 ? null : [lo + n, Math.min(hi + n, max - 1)]
  return hi + 1 >= at ? [lo, Math.min(hi + n, max - 1)] : [lo, hi]
}

/**
 * The rules of the book's sheets after `n` lines are inserted (`n` > 0) or deleted at `at` of `target` (`t`, the
 * references' transform): the target's ranges grow over inserted lines that follow a line of theirs and move with the
 * lines after, shrink with deleted ones, and a rule with none left is gone; every rule's formulas are rewritten. Lines
 * deleted from a colour scale or a data bar (an icon set too, as far as known) put the rule after the others, as Excel
 * puts it back (the set `cond/struct-held`).
 */
export function shiftConditional(sheets: SheetInput[], target: string, axis: Axis, at: number, n: number, t: Transform, names: Map<string, string>): SheetInput[] {
  const max = axis === 'r' ? MAX_ROWS : MAX_COLS
  const count = Math.abs(n)
  // a line of the sheet after the operation where it was before
  const back = (i: number) => (n > 0 ? (i >= at + count ? i - count : i) : i >= at ? i + count : i)
  return sheets.map((s) => {
    const list = s.conditionalFormats
    if (!list?.length) return s
    const own = sheetKey(s.name) === sheetKey(target)
    let changed = false
    const out: { rule: ConditionalFormat; late: boolean }[] = []
    for (const rule of list) {
      const next: Rect[] = []
      let cut = false
      for (const x of own ? rectsOf(rule) : []) {
        const [lo, hi] = axis === 'r' ? [x.r1, x.r2] : [x.c1, x.c2]
        const lohi = n > 0 ? grown(lo, hi, at, count, max) : deleteSpan(lo, hi, at, count, max, false)
        cut ||= n < 0 && lo < at + count && hi >= at
        if (lohi) next.push(axis === 'r' ? { ...x, r1: lohi[0], r2: lohi[1] } : { ...x, c1: lohi[0], c2: lohi[1] })
      }
      if (own && !next.length) {
        changed = true
        continue
      }
      const top = own ? corner(next) : originOf(rule)
      const from = !own ? top : axis === 'r' ? { r: back(top.r), c: top.c } : { r: top.r, c: back(top.c) }
      const p = placed(rule, own ? next : rectsOf(rule), from, s.name, s.name, t, names)
      changed ||= p.rule !== rule
      out.push({ rule: p.rule, late: n < 0 && cut && GRADED.has(rule.type) })
    }
    return changed ? withRules(s, lastLate(out)) : s
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
 * cells take their rules along (to another sheet as rules of their own there, first), the cells they land on keep
 * theirs too, and every rule's formulas are rewritten; a rule whose formulas the move rewrote goes after the others,
 * as Excel puts it back (the set `cond/struct-held`).
 */
export function moveConditional(old: BookInput, out: BookInput, p: Paste, t: Transform, names: Map<string, string>) {
  const [fk, tk] = [sheetKey(p.from), sheetKey(p.to)]
  const [dr, dc] = p.offsets[0]
  const carried: ConditionalFormat[] = []
  out.sheets = out.sheets.map((s) => {
    const k = sheetKey(s.name)
    const list = sheetOf(old, s.name)?.conditionalFormats
    if (!list?.length) return s
    const rules: { rule: ConditionalFormat; late: boolean }[] = []
    for (const rule of list) {
      const rects = rectsOf(rule)
      const stay = k === fk ? subtract(rects, p.src) : rects
      const moved = k === fk ? rects.flatMap((x) => inside(x, p.src) ?? []).map((x) => shifted(x, dr, dc)) : []
      const back = (c: Cell) => (moved.some((x) => c.r >= x.r1 && c.r <= x.r2 && c.c >= x.c1 && c.c <= x.c2) ? { r: c.r - dr, c: c.c - dc } : c)
      const next = fk === tk ? [...stay, ...moved] : stay
      if (next.length) {
        const x = placed(rule, next, back(corner(next)), s.name, s.name, t, names)
        rules.push({ rule: x.rule, late: x.rewrote })
      }
      if (fk !== tk && moved.length) carried.push(placed(rule, moved, { r: corner(moved).r - dr, c: corner(moved).c - dc }, p.from, p.to, t, names).rule)
    }
    return withRules(s, lastLate(rules))
  })
  if (carried.length) out.sheets = out.sheets.map((s) => (sheetKey(s.name) === tk ? withRules(s, [...carried, ...(s.conditionalFormats ?? [])]) : s))
}

/** Whether rectangles together make one rectangle. */
function oneRect(xs: Rect[]): Rect | null {
  const box = { r1: Math.min(...xs.map((x) => x.r1)), r2: Math.max(...xs.map((x) => x.r2)), c1: Math.min(...xs.map((x) => x.c1)), c2: Math.max(...xs.map((x) => x.c2)) }
  const area = (x: Rect) => (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)
  const overlap = xs.some((x, i) => xs.some((y, j) => j > i && meets(x, y)))
  return !overlap && xs.reduce((a, x) => a + area(x), 0) === area(box) ? box : null
}

/**
 * The rules after a copy (`p`, copy and paste or fill; `source`, the book copied from): the cells pasted on lose their
 * rules; the copied cells' rules come as new rules over the pasted cells, first in the target's order, their formulas
 * copied as the cells' are; a rule of the same sheet whose cells and the pasted ones make one rectangle (a fill) grows
 * over them instead (the set `cond/struct-held`).
 */
export function copyConditional(source: BookInput, out: BookInput, p: Paste, names: Map<string, string>) {
  const [fk, tk] = [sheetKey(p.from), sheetKey(p.to)]
  const target = sheetOf(out, p.to)
  const from = sheetOf(source, p.from).conditionalFormats ?? []
  const pasted = (rule: ConditionalFormat) => p.offsets.flatMap(([dr, dc]) => rectsOf(rule).flatMap((x) => inside(x, p.src) ?? []).map((x) => shifted(x, dr, dc)))
  const own = target.conditionalFormats ?? []
  if (!own.length && !from.some((r) => rectsOf(r).some((x) => meets(x, p.src)))) return
  const kept = own.map((rule) => ({ rule, stay: subtract(rectsOf(rule), p.area) }))
  const added: ConditionalFormat[] = []
  for (const rule of from) {
    const next = pasted(rule)
    if (!next.length) continue
    // a fill on the rule's own sheet: one rectangle with the rule's other cells
    const same = fk === tk ? kept.find((x) => x.rule === rule) : undefined
    if (same && oneRect([...same.stay, ...next])) {
      same.stay = [oneRect([...same.stay, ...next])!]
      continue
    }
    const [o, to] = [originOf(rule), corner(next)]
    const copy = copyTransform(to.r - o.r, to.c - o.c)
    added.push({ ...mapFormulas(rule, (text) => rewriteFormula(text, p.to, p.to, copy, names)), ranges: rangesOf(next, []) })
  }
  const list = [...added, ...kept.flatMap(({ rule, stay }) => (stay.length ? [rangesOf(stay, rule.ranges) === rule.ranges ? rule : { ...rule, ranges: rangesOf(stay, rule.ranges) }] : []))]
  out.sheets = out.sheets.map((s) => (s === target ? withRules(s, list) : s))
}
