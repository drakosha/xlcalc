/**
 * Held out from the rows sets rotation and rotation-2, predicted by the rules they gave before Excel ran it: angles and
 * text widths not measured (odd widths, where a projection rounds at a half), sizes between those measured, turned
 * text of other scripts, numbers by formats and booleans turned, turned cells beside level ones of larger fonts and a
 * row's own style, a turned cell merged down, and text typed into a turned cell of a larger font.
 */
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase, size, WRAP } from './cases.js'
import { turned } from './rotation.js'

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  for (const a of [7, 12, 14, 16, 33, 47, 49, 52, 77, -7, -33, -52]) for (const w of [23, 77]) add('angle', { cells: one(textOf(w), turned(a)) })
  for (const a of [30, 60]) for (const w of [13, 27, 61]) add('half', { cells: one(textOf(w), turned(a)) })
  for (const n of [11.5, 17, 22, 26, 30, 40]) for (const a of [30, 75, 90, 'vertical'] as const) add('size', { cells: one('bbbbbbbb', turned(a, size(n))) })
  for (const t of ['Итого', 'Сумма к оплате', 'Ag Ag']) for (const a of [45, 90, 'vertical'] as const) add('script', { cells: one(t, turned(a)) })
  for (const a of [30, 90] as const) {
    add('format', { cells: { 0: { value: 0.1234, format: '0.00%', look: turned(a) } } })
    add('format', { cells: { 0: { value: 1234.5, format: '#,##0.00', look: turned(a) } } })
    add('format', { cells: { 0: { value: 46096, format: 'd mmmm yyyy', look: turned(a) } } })
    add('format', { cells: one(false, turned(a)) })
    add('format', { cells: one(-1234567, turned(a, size(16))) })
  }
  add('cells', { cells: { 0: { value: textOf(30), look: turned(60) }, 1: { value: 'Ag', look: size(36) } } })
  add('cells', { cells: { 0: { value: textOf(30), look: turned(30) }, 1: { value: 'a\nb\nc\nd', look: WRAP }, 2: { value: 'bbb', look: turned('vertical') } } })
  add('cells', { row: { look: size(20) }, cells: one(textOf(20), turned(90)) })
  add('cells', { row: { look: size(8) }, cells: one(textOf(20), turned(45)) })
  add('merged', { cells: one(textOf(60), turned(45)), mergedDown: [0] })
  add('blank', { cells: one('', turned(90, size(36))) })
  add('blank', { cells: one('', turned(30, size(16))) })
  for (const via of ['value', 'keys'] as const) add('typed', { cells: { 0: { look: turned(60, size(16)) } }, typed: [[0, textOf(50)]], via })
  add('typed', { cells: { 0: { look: turned('vertical', size(14)) } }, typed: [[0, 'abcd']], via: 'keys' })
  add('typed', { cells: { 0: { value: textOf(120), look: turned(90) } }, typed: [[0, 'bb']], via: 'value' })
  return out
}
