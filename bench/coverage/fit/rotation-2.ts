/**
 * More of AutoFit of a column under rotated text, to fit the rules `rotation` left open: Calibri of many sizes turned by
 * 90 and 45 degrees and its letters one above the other, the angle every five degrees, numbers in General turned.
 */
import type { CellInput } from '../../../src/contract.js'
import { SIZES, STEPS } from '../rows/rotation-2.js'
import { turned } from '../rows/rotation.js'
import { textOf } from './autofit.js'
import { collector, type FitCase, type FitCell } from './cases.js'

const one = (value: CellInput['value'], look: FitCell['look']): Record<number, FitCell> => ({ 1: { value, look } })

export function buildFitCases(): FitCase[] {
  const { out, add } = collector()
  for (const a of [90, 45, 'vertical'] as const) for (const n of SIZES) add(`s${a === 'vertical' ? 'v' : a}`, { cells: one('bbbbbb', turned(a, { font: { size: n } })) })
  for (const a of STEPS) for (const w of [46, 120]) add('step', { cells: one(textOf(w), turned(a)) })
  for (const a of [15, 30, 45, 60, 75]) for (const v of [1234567890, 12, -5]) add('general', { cells: one(v, turned(a)) })
  return out
}
