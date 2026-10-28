/**
 * Rules of data validation (`SheetInput.validations`): a rule's range as rectangles (xlsx `sqref`, ranges separated by
 * spaces), the rule over a cell, the cell its formulas are written for, the types that compare with bounds.
 */
import type { DataValidation } from '../contract.js'
import { areaText, cornerOf, hasCell, parseRange, type Rect } from '../structure/area.js'

/** The rectangles of a rule's range in order (`B2:B9 D2`), or null when a part is no range of the sheet. */
export function areasOf(range: string): Rect[] | null {
  const parts = range.trim().split(/\s+/)
  const out: Rect[] = []
  for (const p of parts) {
    const x = parseRange(p)
    if (!x) return null
    out.push(x)
  }
  return out.length ? out : null
}

export const rangeText = (xs: Rect[]): string => xs.map(areaText).join(' ')

/** Each rule's rectangles as `ruleAt` last read them, by the range they were read from: a lookup per cell need not parse every rule's range again. */
const read = new WeakMap<DataValidation, { range: string; rects: Rect[] | null }>()

/** The rule over a cell, by its index among the sheet's rules, or -1. */
export function ruleAt(rules: DataValidation[] | undefined, r: number, c: number): number {
  if (!rules) return -1
  return rules.findIndex((v) => {
    let x = read.get(v)
    if (x?.range !== v.range) read.set(v, (x = { range: v.range, rects: areasOf(v.range) }))
    return x.rects?.some((y) => hasCell(y, r, c))
  })
}

/**
 * The cell a rule's formulas are written for: the top-left of the rectangle around its ranges, whatever their order, as
 * for a conditional format (the set `valid/anchor-held`: `$A3>2` over `H3:H6 J1:J3` read from a file is computed for H1,
 * and Excel saves the rule it makes with H3 active as `$A1>2`).
 */
export const anchorOf = (rule: DataValidation): { r: number; c: number } => cornerOf(areasOf(rule.range)!)

/** The types whose value compares with the formulas by `operator`. */
export const COMPARED = new Set(['whole', 'decimal', 'date', 'time', 'textLength'])

/** Whether a rule takes `formula2`: a compared type with `between` (the default) or `notBetween`. */
export const twoBounds = (v: DataValidation): boolean => COMPARED.has(v.type) && (v.operator === undefined || v.operator === 'between' || v.operator === 'notBetween')
