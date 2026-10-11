/**
 * Data validation under structure operations. A rule's ranges move, grow and shrink with inserted and deleted lines as a
 * reference does, a range losing all its lines is gone, and so is a rule losing all its ranges; inserted lines take the
 * rules of the line before them. A move takes the rules of its cells along, the cells it lands on keeping theirs where the
 * moved cells have none; a copy pastes them on every tile, the cells it lands on losing theirs (a pasted cell without a
 * rule clears the target's), as Excel's Paste All does (the set `struct-held`). References in a rule's formulas change as a cell's do, the
 * formulas staying written for the rule's first cell: a rule whose first cell moves or goes has its formulas written
 * again for its new first cell.
 */
import type { BookInput, DataValidation } from '../contract.js'
import { hi, intersect, LIMIT, lo, meets, shiftRect, subtract, withBounds, type Axis, type Rect } from '../structure/area.js'
import { copyTransform } from '../structure/copy.js'
import type { Paste } from '../structure/paste.js'
import { rewriteFormula, sheetOf, type Transform } from '../structure/rewrite.js'
import { deleteSpan, insertSpan } from '../structure/shift.js'
import { sheetKey } from '../syntax.js'
import { anchorOf, areasOf, rangeText } from './rules.js'

type Cell = { r: number; c: number }

/** A rule's formulas, each through `f`; the same rule when none changes. */
function withFormulas(v: DataValidation, f: (text: string) => string): DataValidation {
  const [a, b] = [v.formula1 === undefined ? undefined : f(v.formula1), v.formula2 === undefined ? undefined : f(v.formula2)]
  return a === v.formula1 && b === v.formula2 ? v : { ...v, ...(a !== undefined ? { formula1: a } : {}), ...(b !== undefined ? { formula2: b } : {}) }
}

/** A formula written for cell `from` of `sheet` as at cell `to`: relative references move by the offset, as a copy moves them. */
export function formulaAt(text: string, sheet: string, from: Cell, to: Cell, sheets: Map<string, string>): string {
  return from.r === to.r && from.c === to.c ? text : rewriteFormula(text, sheet, sheet, copyTransform(to.r - from.r, to.c - from.c), sheets)
}

/**
 * The rule over the ranges `areas`: its formulas as at the cell that was `old` before the operation (the new first
 * cell's), then through `t` as a formula of that cell (`host` its sheet before, `moved` after); null without ranges.
 */
function placed(v: DataValidation, areas: Rect[], sheet: string, old: Cell, host: string, moved: string, t: Transform | null, sheets: Map<string, string>): DataValidation | null {
  if (!areas.length) return null
  const anchor = anchorOf(v)
  const at = withFormulas(v, (text) => {
    const there = formulaAt(text, sheet, anchor, old, sheets)
    return t ? rewriteFormula(there, host, moved, t, sheets) : there
  })
  const range = rangeText(areas)
  return range === v.range ? at : { ...at, range }
}

/** The references of the rules' formulas through `t` (a sheet's rules, `host` its name before, `moved` after); the same array when none changes. */
export function rewriteValidations(rules: DataValidation[] | undefined, host: string, moved: string, t: Transform, sheets: Map<string, string>): DataValidation[] | undefined {
  if (!rules) return rules
  let out: DataValidation[] | undefined
  rules.forEach((v, i) => {
    const next = withFormulas(v, (text) => rewriteFormula(text, host, moved, t, sheets))
    if (next !== v) (out ??= [...rules])[i] = next
  })
  return out ?? rules
}

/**
 * The rules of `sheet` after `n` lines inserted before `at` (n > 0) or deleted from `at` (n < 0) along `axis`; `t` the
 * operation's transform. A range follows the references' rules for ranges (`insertSpan`, `deleteSpan`): it grows with
 * lines inserted after its first line, moves with lines inserted at or before it, and leaves the sheet's end in place, a
 * range of one cell too (the set `edge-held`: B1048576 is B1048575:B1048576 after a row above is deleted); and it takes
 * in the lines inserted right after it, which take the rules of the line before them (the set `struct-held`).
 */
export function shiftValidations(rules: DataValidation[] | undefined, sheet: string, axis: Axis, at: number, n: number, t: Transform, sheets: Map<string, string>): DataValidation[] | undefined {
  if (!rules) return rules
  const count = Math.abs(n)
  const back = (i: number) => (i < at ? i : n > 0 ? i - count : i + count)
  const out: DataValidation[] = []
  let changed = false
  for (const v of rules) {
    const areas = areasOf(v.range)!.flatMap((x) => {
      const [a, b] = [lo(x, axis), hi(x, axis)]
      const one = a === b && (axis === 'r' ? x.c1 === x.c2 : x.r1 === x.r2)
      // a range whose last line is the one before the inserted lines takes them in, as their cells take that line's look
      const s = n > 0 ? (b === at - 1 ? [a, Math.min(b + count, LIMIT[axis] - 1)] : insertSpan(a, b, at, count, LIMIT[axis], one)) : deleteSpan(a, b, at, count, LIMIT[axis], false)
      return s ? [withBounds(x, axis, s[0], s[1])] : []
    })
    const first = areas[0]
    const old = first && (axis === 'r' ? { r: back(first.r1), c: first.c1 } : { r: first.r1, c: back(first.c1) })
    const next = first ? placed(v, areas, sheet, old!, sheet, sheet, t, sheets) : null
    changed ||= next !== v
    if (next) out.push(next)
  }
  if (!changed) return rules
  return out.length ? out : undefined
}

/** Rectangles side by side over the same rows, or one above the other over the same columns, joined; the first stays first. */
function coalesce(xs: Rect[]): Rect[] {
  const out = [...xs]
  for (let i = 0; i < out.length; i++)
    for (let j = i + 1; j < out.length; j++) {
      const [a, b] = [out[i], out[j]]
      const rows = a.c1 === b.c1 && a.c2 === b.c2 && (a.r2 + 1 === b.r1 || b.r2 + 1 === a.r1)
      const cols = a.r1 === b.r1 && a.r2 === b.r2 && (a.c2 + 1 === b.c1 || b.c2 + 1 === a.c1)
      if (!rows && !cols) continue
      out[i] = { r1: Math.min(a.r1, b.r1), r2: Math.max(a.r2, b.r2), c1: Math.min(a.c1, b.c1), c2: Math.max(a.c2, b.c2) }
      out.splice(j, 1)
      j = i
    }
  return out
}

/** A rule with its ranges joined where they touch, its formulas written for its first cell when that moves up or left. */
function tidy(v: DataValidation, sheet: string, sheets: Map<string, string>): DataValidation {
  const xs = areasOf(v.range)!
  const ys = coalesce(xs)
  if (ys.length === xs.length) return v
  const [from, to] = [anchorOf(v), { r: ys[0].r1, c: ys[0].c1 }]
  return { ...withFormulas(v, (text) => formulaAt(text, sheet, from, to, sheets)), range: rangeText(ys) }
}

/** The settings of a rule but its range and formulas, to find rules that are one. */
const settings = ({ range: _, formula1: __, formula2: ___, ...rest }: DataValidation) => JSON.stringify(Object.entries(rest).sort(([a], [b]) => (a < b ? -1 : 1)))

/** Rules of a sheet joined where they are one rule: the same settings, and formulas that are the first one's at the other's first cell. */
function joined(rules: DataValidation[], sheet: string, sheets: Map<string, string>): DataValidation[] {
  const out: DataValidation[] = []
  for (const v of rules) {
    const same = out.findIndex((w) => {
      if (settings(w) !== settings(v)) return false
      const [a, b] = [anchorOf(w), anchorOf(v)]
      return (['formula1', 'formula2'] as const).every((k) => (w[k] === undefined ? v[k] === undefined : v[k] !== undefined && formulaAt(w[k], sheet, a, b, sheets) === v[k]))
    })
    if (same < 0) out.push(v)
    else out[same] = { ...out[same], range: `${out[same].range} ${v.range}` }
  }
  return out
}

/** Rules without the cells of `holes`: a rule whose first cell goes has its formulas written for its new first cell, one left without cells goes. */
function withoutAreas(rules: DataValidation[], holes: Rect[], sheet: string, sheets: Map<string, string>): DataValidation[] {
  const kept: DataValidation[] = []
  for (const v of rules) {
    const areas = areasOf(v.range)!.flatMap((x) => holes.reduce(subtract, [x]))
    const first = areas[0]
    const next = first ? placed(v, areas, sheet, { r: first.r1, c: first.c1 }, sheet, sheet, null, sheets) : null
    if (next) kept.push(next)
  }
  return kept
}

/** A sheet's rules without the cells of `holes` (Excel's Clear over them), their ranges joined where they touch; the same array when no rule meets them. */
export function clearValidations(rules: DataValidation[] | undefined, holes: Rect[], sheet: string, sheets: Map<string, string>): DataValidation[] | undefined {
  if (!rules?.some((v) => areasOf(v.range)!.some((x) => holes.some((h) => meets(x, h))))) return rules
  const kept = withoutAreas(rules, holes, sheet, sheets).map((v) => tidy(v, sheet, sheets))
  return kept.length ? kept : undefined
}

/**
 * The rules after a move or a copy (`p`), in `out` (its sheets replaced where their rules change): the cells the paste
 * lands on and, for a move, the cells it takes lose their rules; the source's rules (of `source`, the book as it was
 * copied) are pasted at each offset, their formulas as a move (`t`) or a copy moves a formula of their first pasted cell.
 */
export function pasteValidations(source: BookInput, out: BookInput, p: Paste, t: Transform | null, sheets: Map<string, string>) {
  const [fk, tk] = [sheetKey(p.from), sheetKey(p.to)]
  const src = sheetOf(source, p.from)?.validations ?? []
  const pasted: DataValidation[] = []
  // where a move lands, the cells its rules land on
  const landed: Rect[] = []
  for (const [dr, dc] of p.offsets)
    for (const v of src) {
      const pieces = areasOf(v.range)!.flatMap((x) => intersect(x, p.src) ?? [])
      if (!pieces.length) continue
      const areas = pieces.map((x) => shiftRect(x, dr, dc))
      landed.push(...areas)
      // a move's formulas as a moved cell's (`t`), a copy's and a move's to another sheet as a copied cell's (the set `struct-held-2`)
      const moved = p.move && fk === tk
      const next = placed(v, areas, p.from, { r: pieces[0].r1, c: pieces[0].c1 }, moved ? p.from : p.to, p.to, moved ? t : copyTransform(dr, dc), sheets)
      if (next) pasted.push(next)
    }
  const gone = (k: string) => [...(k === tk ? (p.move ? landed : [p.area]) : []), ...(p.move && k === fk ? [p.src] : [])]
  out.sheets = out.sheets.map((s) => {
    const k = sheetKey(s.name)
    const holes = gone(k)
    let rules = s.validations
    if (rules && holes.length) rules = withoutAreas(rules, holes, s.name, sheets)
    if (t && rules) rules = rewriteValidations(rules, s.name, s.name, t, sheets)
    if (k === tk && pasted.length) rules = joined([...(rules ?? []), ...pasted], s.name, sheets)
    if (rules === s.validations) return s
    rules = rules?.map((v) => tidy(v, s.name, sheets))
    const { validations: _, ...rest } = s
    return rules?.length ? { ...rest, validations: rules } : rest
  })
}
