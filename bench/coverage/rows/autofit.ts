/**
 * AutoFit of a row: an empty row, Calibri of every size from 1 to 409 points, bold and italic, other fonts, values of
 * every kind, several cells and the row's or column's style, text with line breaks wrapped and not, where a wrapped line
 * breaks (two words and one long word about the default column's room), long text over columns of several widths, a
 * height set by hand and a hidden row, merged cells, a formula's text; and text with line breaks set through the object
 * model, which may or may not wrap it.
 */
import type { CellStyle } from '../../../src/contract.js'
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase, size, WRAP } from './cases.js'

const SIZES = [1, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 96, 144, 200, 300, 409]

const FONTS: CellStyle['font'][] = [
  { bold: true },
  { italic: true },
  { bold: true, italic: true, size: 16 },
  { underline: 'double' },
  { name: 'Arial', size: 10 },
  { name: 'Arial', size: 11 },
  { name: 'Arial', size: 16 },
  { name: 'Times New Roman', size: 11 },
  { name: 'Times New Roman', size: 12 },
  { name: 'Courier New', size: 10 },
  { name: 'Cambria', size: 11 },
  { name: 'Verdana', size: 10 },
  { name: 'Helvetica', size: 12 },
  { name: 'Georgia', size: 11 },
  { name: 'Calibri Light', size: 11 },
  { name: 'Aptos', size: 11 },
  { name: 'Aptos Narrow', size: 11 },
  { name: 'No Such Font', size: 11 },
]

const SENTENCE = 'The quick brown fox jumps over the lazy dog. '.repeat(3).trim()

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  // nothing to show: an empty row, one set by hand, a large font on a blank, empty text, a blank wrapped
  add('empty', {})
  add('empty', { row: { height: 30 } })
  add('empty', { cells: { 0: { look: size(16) } } })
  add('empty', { cells: one('', size(16)) })
  add('empty', { cells: { 0: { look: WRAP } } })
  // Calibri by its size, and other fonts
  for (const n of SIZES) add('size', { cells: one('Ag', size(n)) })
  for (const font of FONTS) add('font', { cells: one('Ag', { font }) })
  // values other than text, in a larger font
  add('value', { cells: one(123, size(16)) })
  add('value', { cells: one(true, size(16)) })
  add('value', { cells: { 0: { formula: '1/0', look: size(16) } } })
  add('value', { cells: { 0: { value: 46096, format: 14, look: size(16) } } })
  // several cells: the highest counts; the row's style, a column's style
  add('many', { cells: { 0: { value: 'a' }, 1: { value: 'b', look: size(16) }, 2: { value: 'c', look: size(24) } } })
  add('many', { row: { look: size(16) } })
  add('many', { row: { look: size(16) }, cells: one('a') })
  add('many', { row: { look: size(16) }, cells: one('a', size(11)) })
  add('many', { columns: { 1: { look: size(20) } }, cells: { 1: { value: 'a' } } })
  add('many', { columns: { 1: { look: size(20) } } })
  // line breaks, not wrapped and wrapped
  add('lines', { cells: one('a\nb') })
  for (const t of ['a\nb', 'a\nb\nc', 'a\nb\nc\nd\ne', Array.from({ length: 10 }, (_, i) => `l${i}`).join('\n'), 'a\n', '\n', 'a\n\nb', '\na']) add('lines', { cells: one(t, WRAP) })
  add('lines', { cells: one('a\nb', size(16, WRAP)) })
  add('lines', { cells: one('a\nb', { font: { name: 'Arial', size: 10 }, alignment: { wrap: true } }) })
  add('lines', { cells: one(Array.from({ length: 30 }, (_, i) => `l${i}`).join('\n'), WRAP) })
  // where a wrapped line breaks in the default column (53 points): a word and another of w points, and one long word
  for (let w = 34; w <= 50; w++) add('room', { cells: one(`b ${textOf(w)}`, WRAP) })
  for (let w = 42; w <= 54; w++) add('word', { cells: one(textOf(w), WRAP) })
  // long text over columns of several widths, a larger font, Cyrillic
  for (const width of [3, 5, 10, 20, 40, 80]) add('width', { cells: one(SENTENCE, WRAP), columns: { 0: { width } } })
  add('width', { cells: one(SENTENCE, size(16, WRAP)), columns: { 0: { width: 20 } } })
  add('width', { cells: one('Съешь же ещё этих мягких французских булок, да выпей чаю.', WRAP), columns: { 0: { width: 10 } } })
  add('width', { cells: one(SENTENCE, WRAP), columns: { 0: { width: 10, hidden: true } } })
  // a height set by hand, a hidden row
  add('hand', { row: { height: 30 }, cells: one('a') })
  add('hand', { row: { height: 30 }, cells: one('a\nb\nc', WRAP) })
  add('hand', { row: { height: 5 }, cells: one('a', size(16)) })
  add('hand', { row: { hidden: true }, cells: one('a\nb', WRAP) })
  // merged cells: across into the next column, down into the next row
  add('merged', { cells: one('a\nb\nc', WRAP), merged: [0] })
  add('merged', { cells: one('Ag', size(24)), merged: [0] })
  add('merged', { cells: one('Ag', size(24)), mergedDown: [0] })
  add('merged', { cells: { 0: { value: 'a\nb\nc', look: WRAP }, 2: { value: 'a\nb', look: WRAP } }, merged: [0] })
  // a number too wide for its wrapped column, a formula's text with a line break
  add('number', { cells: one(123456789, WRAP), columns: { 0: { width: 3 } } })
  add('number', { cells: one(1234.5, { ...WRAP, font: { size: 16 } }), columns: { 0: { width: 3 } } })
  add('formula', { cells: { 0: { formula: '"a"&CHAR(10)&"b"', look: WRAP } } })
  add('formula', { cells: { 0: { formula: '"a"&CHAR(10)&"b"' } } })
  // line breaks set through the object model: whether the cell wraps and the row grows, then AutoFit
  for (const via of ['value', 'local'] as const) {
    add('typed', { typed: [[0, 'a\nb']], via })
    add('typed', { typed: [[0, 'a\nb']], via, fitAfter: true })
    add('typed', { row: { height: 30 }, typed: [[0, 'a\nb\nc']], via })
    add('typed', { cells: { 0: { look: WRAP } }, typed: [[0, 'a\nb\nc']], via })
  }
  return out
}
