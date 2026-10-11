/**
 * Rules of data validation (`SheetInput.validations`): a rule's range as rectangles (xlsx `sqref`, ranges separated by
 * spaces), the rule over a cell, the cell its formulas are written for, the types that compare with bounds.
 */
import type { DataValidation } from '../contract.js'
import { areaText, hasCell, parseRange, type Rect } from '../structure/area.js'

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

/** The rule over a cell, by its index among the sheet's rules, or -1. */
export function ruleAt(rules: DataValidation[] | undefined, r: number, c: number): number {
  if (!rules) return -1
  return rules.findIndex((v) => areasOf(v.range)?.some((x) => hasCell(x, r, c)))
}

/** The cell a rule's formulas are written for: the top-left cell of its first range. */
export function anchorOf(rule: DataValidation): { r: number; c: number } {
  const x = areasOf(rule.range)![0]
  return { r: x.r1, c: x.c1 }
}

/** The types whose value compares with the formulas by `operator`. */
export const COMPARED = new Set(['whole', 'decimal', 'date', 'time', 'textLength'])

/** Whether a rule takes `formula2`: a compared type with `between` (the default) or `notBetween`. */
export const twoBounds = (v: DataValidation): boolean => COMPARED.has(v.type) && (v.operator === undefined || v.operator === 'between' || v.operator === 'notBetween')
