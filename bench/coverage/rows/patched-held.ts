/**
 * Held out from the rows set patched: a rotation by an angle and a smaller font set on whole columns, a larger font on
 * columns where another column holds the row's text, and the sheet's default height as Excel saves it after a font
 * on whole columns.
 */
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase } from './cases.js'

const MORE = [1, 3, 200, 1048576]

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  add('rotation', { cells: one(textOf(60)), patched: { cols: [0, 0], rotation: 45 }, more: MORE })
  add('rotation', { cells: one(textOf(30)), patched: { cols: [0, 0], rotation: -90 }, more: MORE })
  add('size', { cells: one('a'), patched: { cols: [0, 0], size: 8 }, more: MORE })
  add('size', { cells: { 2: { value: 'a' } }, patched: { cols: [0, 0], size: 28 }, more: MORE })
  add('both', { cells: one('bbbbbb'), patched: { cols: [0, 0], size: 16, rotation: 90 }, more: MORE })
  return out
}
