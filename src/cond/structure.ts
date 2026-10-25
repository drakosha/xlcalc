/**
 * Conditional formats under structure operations, as Excel changes them (the set `cond/struct-held`). A rule's ranges grow
 * and shrink with inserted and deleted lines, go with moved cells and are copied with copied ones. Its formulas change as
 * the formula of a cell at the rule's first cell (the top-left of the rectangle around its ranges) would, and when that
 * cell is no longer first they are made for the new first cell as a copy would make them. Excel also changes the order
 * of rules: some it takes out and puts back after the others, and pasted rules come first.
 */
import type { BookInput, ConditionalFormat, ConditionalValue, SheetInput } from '../contract.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { sheetKey, tokenize } from '../syntax.js'
import { cornerOf, hasCell, hi, intersect, lo, meets, shiftRect, subtract, withBounds, type Axis, type Rect } from '../structure/area.js'
import type { Paste } from '../structure/paste.js'
import { rewriteFormula, sheetOf, type Transform } from '../structure/rewrite.js'
import { copyTransform, deleteSpan, growSpan } from '../structure/transform.js'
import { originOf, rangesOf, rectsOf } from './rules.js'

type Cell = { r: number; c: number }


const valueWith = (v: ConditionalValue, f: (text: string) => string): ConditionalValue => (v.value === undefined ? v : { ...v, value: f(v.value) })

/** The rule with each of its formulas through `f`. */
export function mapFormulas(rule: ConditionalFormat, f: (text: string) => string): ConditionalFormat {
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
 * A rule over `next` (rectangles of the sheet after the operation), its formulas made for its new first cell (`cornerOf`):
 * copied from the old one to `from`, that cell's place before the operation, then rewritten by `t`.
 */
function placed(rule: ConditionalFormat, next: Rect[], from: Cell, host: string, moved: string, t: Transform, names: Map<string, string>): Placed {
  const o = originOf(rule)
  const [dr, dc] = [from.r - o.r, from.c - o.c]
  // the formula as at another cell of the rule, wrapping round the sheet's edges as it computes
  const copy = dr || dc ? copyTransform(dr, dc, true) : null
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

/** Rules over the statistics of their cells. */
const GRADED = new Set(['colorScale', 'dataBar', 'iconSet', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues'])

/**
 * The rules of the book's sheets after `n` lines are inserted (`n` > 0) or deleted at `at` of `target` (`t`, the
 * references' transform): the target's ranges grow over inserted lines that follow a line of theirs and move with the
 * lines after, shrink with deleted ones, and a rule with none left is gone; every rule's formulas are rewritten. Lines
 * deleted from a rule over the statistics of its cells (a colour scale, a data bar, an icon set, a top rule; average and
 * duplicate rules taken as these) put the rule after the others, as Excel puts it back (the sets `cond/struct-held`, `held`).
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
        const [a, b] = [lo(x, axis), hi(x, axis)]
        const lohi = n > 0 ? growSpan(a, b, at, count, max, false) : deleteSpan(a, b, at, count, max, false)
        cut ||= n < 0 && a < at + count && b >= at
        if (lohi) next.push(withBounds(x, axis, lohi[0], lohi[1]))
      }
      if (own && !next.length) {
        changed = true
        continue
      }
      const top = own ? cornerOf(next) : originOf(rule)
      const from = !own ? top : axis === 'r' ? { r: back(top.r), c: top.c } : { r: top.r, c: back(top.c) }
      const p = placed(rule, own ? next : rectsOf(rule), from, s.name, s.name, t, names)
      changed ||= p.rule !== rule
      out.push({ rule: p.rule, late: n < 0 && cut && GRADED.has(rule.type) })
    }
    return changed ? withRules(s, lastLate(out)) : s
  })
}

/** A transform that rewrites no reference. */
const KEEP: Transform = { ref: () => undefined }

/**
 * A sheet's rules without the cells of `holes` (Excel's ClearFormats over them): a rule left without cells goes, and one
 * whose first cell goes has its formulas made for its new first cell; the same list when no rule meets them.
 */
export function clearConditional(list: ConditionalFormat[] | undefined, holes: Rect[], host: string, names: Map<string, string>): ConditionalFormat[] | undefined {
  if (!list?.some((rule) => rectsOf(rule).some((x) => holes.some((h) => meets(x, h))))) return list
  const out = list.flatMap((rule) => {
    const stay = holes.reduce((xs, h) => subtract(xs, h), rectsOf(rule))
    return stay.length ? [placed(rule, stay, cornerOf(stay), host, host, KEEP, names).rule] : []
  })
  return out.length ? out : undefined
}

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
      const moved = k === fk ? rects.flatMap((x) => intersect(x, p.src) ?? []).map((x) => shiftRect(x, dr, dc)) : []
      const back = (c: Cell) => (moved.some((x) => hasCell(x, c.r, c.c)) ? { r: c.r - dr, c: c.c - dc } : c)
      const next = fk === tk ? [...stay, ...moved] : stay
      if (next.length) {
        const x = placed(rule, next, back(cornerOf(next)), s.name, s.name, t, names)
        rules.push({ rule: x.rule, late: x.rewrote })
      }
      if (fk !== tk && moved.length) carried.push(placed(rule, moved, { r: cornerOf(moved).r - dr, c: cornerOf(moved).c - dc }, p.from, p.to, t, names).rule)
    }
    return withRules(s, lastLate(rules))
  })
  if (carried.length) out.sheets = out.sheets.map((s) => (sheetKey(s.name) === tk ? withRules(s, [...carried, ...(s.conditionalFormats ?? [])]) : s))
}

/**
 * The one rectangle rectangles of the same columns make, stacked without gaps or overlaps: a rule grows over a paste
 * that goes on down or up its columns (the sets `cond/struct-held`, `held`: a fill down grows it, a copy right next to
 * it is a new rule).
 */
function column(xs: Rect[]): Rect | null {
  if (xs.some((x) => x.c1 !== xs[0].c1 || x.c2 !== xs[0].c2)) return null
  const ys = [...xs].sort((a, b) => a.r1 - b.r1)
  for (let i = 1; i < ys.length; i++) if (ys[i].r1 !== ys[i - 1].r2 + 1) return null
  return { r1: ys[0].r1, r2: ys[ys.length - 1].r2, c1: ys[0].c1, c2: ys[0].c2 }
}

/**
 * The rules after a copy (`p`, copy and paste or fill; `source`, the book copied from): the cells pasted on lose their
 * rules; the copied cells' rules come as new rules over the pasted cells, first in the target's order, their formulas
 * copied as the cells' are; a rule of the same sheet whose cells and the pasted ones make one rectangle down its columns
 * (a fill down) grows over them instead (the sets `cond/struct-held`, `held`). A rule whose first cell moves has its
 * formulas made for the new one.
 */
export function copyConditional(source: BookInput, out: BookInput, p: Paste, names: Map<string, string>) {
  const [fk, tk] = [sheetKey(p.from), sheetKey(p.to)]
  const target = sheetOf(out, p.to)
  const from = sheetOf(source, p.from).conditionalFormats ?? []
  const pasted = (rule: ConditionalFormat) => p.offsets.flatMap(([dr, dc]) => rectsOf(rule).flatMap((x) => intersect(x, p.src) ?? []).map((x) => shiftRect(x, dr, dc)))
  const own = target.conditionalFormats ?? []
  if (!own.length && !from.some((r) => rectsOf(r).some((x) => meets(x, p.src)))) return
  const kept = own.map((rule) => ({ rule, stay: subtract(rectsOf(rule), p.area) }))
  const added: ConditionalFormat[] = []
  for (const rule of from) {
    const next = pasted(rule)
    if (!next.length) continue
    // a fill down or up on the rule's own sheet: one column of rectangles with the rule's other cells
    const same = fk === tk ? kept.find((x) => x.rule === rule) : undefined
    const grownTo = same && column([...same.stay, ...next])
    if (same && grownTo) {
      same.stay = [grownTo]
      continue
    }
    const [o, to] = [originOf(rule), cornerOf(next)]
    const copy = copyTransform(to.r - o.r, to.c - o.c, true)
    added.push({ ...mapFormulas(rule, (text) => rewriteFormula(text, p.to, p.to, copy, names)), ranges: rangesOf(next, []) })
  }
  // a rule whose first cell the paste took or a fill up moved has its formulas made for the new one, as clearConditional does
  const list = [...added, ...kept.flatMap(({ rule, stay }) => (stay.length ? [placed(rule, stay, cornerOf(stay), p.to, p.to, KEEP, names).rule] : []))]
  out.sheets = out.sheets.map((s) => (s === target ? withRules(s, list) : s))
}
