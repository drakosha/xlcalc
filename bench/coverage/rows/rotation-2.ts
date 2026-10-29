/**
 * More of AutoFit of a row under rotated text, to fit the rules `rotation` left open: Calibri of many sizes turned by
 * 90 and 45 degrees over two lengths of text, the angle every five degrees, numbers in General turned, a blank and empty
 * text of a larger font turned, and wrapped text of a larger font turned.
 */
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase, size, WRAP } from './cases.js'
import { turned } from './rotation.js'

/** Sizes of Calibri measured turned. */
export const SIZES = [6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 18, 20, 24, 28, 36, 48]

/** Angles every five degrees, but those `rotation` measured, and a few near the ends. */
export const STEPS = [2, 3, 5, 10, 20, 25, 35, 40, 50, 55, 65, 70, 80, 85, 87, 88]

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  for (const n of SIZES) for (const t of ['bbbbbb', 'b'.repeat(12)]) add('s90', { cells: one(t, turned(90, size(n))) })
  for (const n of SIZES) for (const t of ['bbbbbb', 'b'.repeat(12)]) add('s45', { cells: one(t, turned(45, size(n))) })
  for (const a of STEPS) for (const w of [46, 120]) add('step', { cells: one(textOf(w), turned(a)) })
  for (const a of [15, 30, 45, 60, 75, 90]) for (const v of [1234567, 12, -5]) add('general', { cells: one(v, turned(a)) })
  for (const a of [45, 'vertical'] as const) {
    add('blank', { cells: { 0: { look: turned(a, size(24)) } } })
    add('blank', { cells: one('', turned(a, size(24))) })
  }
  for (const a of [30, 45, 60] as const) add('wrap', { cells: one('a\nb', turned(a, size(24, WRAP))) })
  return out
}
