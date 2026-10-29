/**
 * AutoFit of a row under rotated text (`alignment.rotation`): angles from -90 to 90 and the letters one above the other
 * over texts of several widths, larger fonts, wrapped text turned, several turned cells and a level one beside them, a
 * number, a date, a boolean and an error turned, merged cells, a blank turned cell; and text typed into a turned cell
 * through the object model and on the keyboard, which may fit the row.
 */
import type { CellStyle } from '../../../src/contract.js'
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase, size, WRAP } from './cases.js'

/** A style turned by `rotation`, over `more`. */
export const turned = (rotation: number | 'vertical', more: CellStyle = {}): CellStyle => ({ ...more, alignment: { ...more.alignment, rotation } })

const ANGLES = [1, 15, 30, 45, 60, 75, 89, 90, -1, -15, -30, -45, -60, -90] as const

const SENTENCE = 'The quick brown fox jumps over the lazy dog. '.repeat(3).trim()

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  // the angle over texts of 12, 46 and 120 points
  for (const a of ANGLES) for (const w of [12, 46, 120]) add('angle', { cells: one(textOf(w), turned(a)) })
  // the letters one above the other: their count, a space, a line break, letters of other heights
  for (const t of ['b', 'bb', 'bbbbb', 'b'.repeat(10), 'b'.repeat(20), 'Ag', 'a b', 'a\nb', 'Wij', '12345']) add('vertical', { cells: one(t, turned('vertical')) })
  // a larger font turned
  for (const n of [8, 16, 24]) for (const a of [45, 90, -90, 'vertical'] as const) add('size', { cells: one('bbbbbb', turned(a, size(n))) })
  // long text: the highest row
  for (const a of [45, 90] as const) add('long', { cells: one(textOf(400), turned(a)) })
  add('long', { cells: one('b'.repeat(40), turned('vertical')) })
  // wrapped and turned: a long text, line breaks
  for (const a of [45, 90, -90, 'vertical'] as const) add('wrap', { cells: one(SENTENCE, turned(a, WRAP)) })
  for (const a of [45, 90, 'vertical'] as const) add('wrap', { cells: one('a\nb', turned(a, WRAP)) })
  for (const a of [45, 90] as const) add('wrap', { cells: one('a\nb', turned(a)) })
  add('wrap', { cells: one(SENTENCE, turned(90, WRAP)), columns: { 0: { width: 20 } } })
  // several cells: two turned, a turned one beside a level larger font or a wrapped one
  add('cells', { cells: { 0: { value: textOf(46), look: turned(45) }, 1: { value: textOf(46), look: turned(90) } } })
  add('cells', { cells: { 0: { value: textOf(46), look: turned(-45) }, 1: { value: textOf(120), look: turned(30) } } })
  add('cells', { cells: { 0: { value: textOf(46), look: turned(90) }, 1: { value: 'Ag', look: size(24) } } })
  add('cells', { cells: { 0: { value: textOf(12), look: turned(90) }, 1: { value: 'a\nb\nc', look: WRAP } } })
  add('cells', { cells: { 0: { value: 'bbbbb', look: turned('vertical') }, 1: { value: textOf(46), look: turned(90) } } })
  // values other than text turned
  for (const a of [45, 90] as const) {
    add('value', { cells: one(1234567, turned(a)) })
    add('value', { cells: { 0: { value: 46096, format: 14, look: turned(a) } } })
    add('value', { cells: one(true, turned(a)) })
    add('value', { cells: { 0: { formula: '1/0', look: turned(a) } } })
  }
  add('value', { cells: one(3.14159265358979, turned(90)) })
  add('value', { cells: one(1234567, turned('vertical')) })
  // merged cells turned: across, down
  add('merged', { cells: one(textOf(46), turned(90)), merged: [0] })
  add('merged', { cells: one(textOf(46), turned(90)), mergedDown: [0] })
  add('merged', { cells: one('bbbbb', turned('vertical')), merged: [0] })
  // nothing to show turned: a blank, empty text
  add('blank', { cells: { 0: { look: turned(90) } } })
  add('blank', { cells: one('', turned(90)) })
  add('blank', { cells: { 0: { look: turned(90, size(24)) } } })
  // a height set by hand, then AutoFit
  add('hand', { row: { height: 30 }, cells: one(textOf(46), turned(90)) })
  // text typed into a turned cell: whether the row fits it, then AutoFit
  for (const via of ['value', 'local', 'keys'] as const) add('typed', { cells: { 0: { look: turned(90) } }, typed: [[0, textOf(46)]], via })
  add('typed', { cells: { 0: { look: turned('vertical') } }, typed: [[0, 'bbbbb']], via: 'keys' })
  add('typed', { cells: { 0: { look: turned(45) } }, typed: [[0, '1234567']], via: 'keys' })
  add('typed', { cells: { 0: { look: turned(90) } }, typed: [[0, textOf(46)]], via: 'value', fitAfter: true })
  add('typed', { row: { height: 30 }, cells: { 0: { look: turned(90) } }, typed: [[0, textOf(46)]], via: 'keys' })
  return out
}
