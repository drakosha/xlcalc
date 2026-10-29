/**
 * AutoFit of a column under rotated text (`alignment.rotation`): angles from -90 to 90 and the letters one above the
 * other over texts of several widths, a larger font, wrapped text turned, a turned cell beside a level one, numbers and
 * dates turned, merged cells, a blank turned cell; and a number typed into a turned cell, which may widen its column.
 */
import type { CellInput } from '../../../src/contract.js'
import { turned } from '../rows/rotation.js'
import { textOf } from './autofit.js'
import { collector, type FitCase, type FitCell } from './cases.js'

const ANGLES = [1, 15, 30, 45, 60, 75, 89, 90, -1, -15, -30, -45, -60, -90] as const

const one = (value: CellInput['value'], look: FitCell['look'], more: Partial<FitCell> = {}): Record<number, FitCell> => ({ 1: { value, look, ...more } })

export function buildFitCases(): FitCase[] {
  const { out, add } = collector()
  // the angle over texts of 12, 46 and 120 points
  for (const a of ANGLES) for (const w of [12, 46, 120]) add('angle', { cells: one(textOf(w), turned(a)) })
  // the letters one above the other: their count, letters of other widths, a line break
  for (const t of ['b', 'bbbbb', 'b'.repeat(20), 'W', 'Wij', 'a b', 'a\nb', '12345']) add('vertical', { cells: one(t, turned('vertical')) })
  // a larger font turned
  for (const n of [8, 16, 24]) for (const a of [45, 90, 'vertical'] as const) add('size', { cells: one('bbbbbb', turned(a, { font: { size: n } })) })
  // wrapped and turned
  for (const a of [45, 90, 'vertical'] as const) add('wrap', { cells: one('The quick brown fox jumps over the lazy dog.', turned(a, { alignment: { wrap: true } })) })
  for (const a of [45, 90] as const) add('wrap', { cells: one('a\nb', turned(a, { alignment: { wrap: true } })) })
  // a turned cell beside a level one, two turned ones
  add('cells', { cells: { 1: { value: textOf(120), look: turned(90) }, 2: { value: textOf(30) } } })
  add('cells', { cells: { 1: { value: textOf(120), look: turned(45) }, 2: { value: textOf(120) } } })
  add('cells', { cells: { 1: { value: textOf(120), look: turned(45) }, 2: { value: textOf(46), look: turned(-30) } } })
  // values other than text turned
  for (const a of [45, 90] as const) {
    add('value', { cells: one(1234567890, turned(a)) })
    add('value', { cells: one(46096, turned(a), { format: 'd mmmm yyyy' }) })
    add('value', { cells: one(true, turned(a)) })
  }
  add('value', { cells: one(1234567890, turned('vertical')) })
  // merged cells turned: across, down
  add('merged', { cells: one(textOf(120), turned(45)), merged: [1] })
  add('merged', { cells: { 1: { value: textOf(120), look: turned(45) } }, mergedDown: [1, 3] })
  // nothing to show turned: a blank, empty text
  add('blank', { cells: { 1: { look: turned(90) } } })
  add('blank', { cells: one('', turned(90)) })
  // a column narrowed by hand, then AutoFit
  add('hand', { column: { width: 2 }, cells: one(textOf(120), turned(45)) })
  // a number typed into a turned cell: whether the column widens
  for (const a of [45, 90] as const) add('typed', { cells: { 1: { look: turned(a) } }, typed: [[1, '1234567890']], via: 'value' })
  add('typed', { cells: { 1: { look: turned(90) } }, typed: [[1, '1234567890']], via: 'keys' })
  add('typed', { cells: { 1: { look: turned(90) } }, typed: [[1, textOf(120)]], via: 'keys' })
  return out
}
