/**
 * Row heights held out of `autofit`, predicted by the rules it gave: Calibri of sizes not measured, wrapped lines of
 * every size, other fonts wrapped, the row's style smaller than the default, a wrapped cell beside a larger font, where a
 * wrapped line of 16 points breaks, the room at another width, spaces and a hyphen at the break, a wrapped boolean and
 * error, a merged cell of a larger font, hidden rows, bold text wrapped, long text over more widths; and through the
 * object model a formula with a line break or giving one, text into a wrapped cell, a line break into a larger font, a
 * merged cell and a row of a taller cell, a carriage return.
 */
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase, size, WRAP } from './cases.js'

const lines = (n: number) => Array.from({ length: n }, (_, i) => `l${i}`).join('\n')

const SENTENCE = 'The quick brown fox jumps over the lazy dog. '.repeat(3).trim()

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  // Calibri of sizes the first set did not measure, halves too
  for (const n of [7, 17, 19, 21, 23, 25, 27, 30, 32, 33, 40, 50, 60, 100, 250, 10.5, 11.5, 12.5, 14.5, 16.5]) add('size', { cells: one('Ag', size(n)) })
  // wrapped lines by size: two lines, one, three; and many lines up to the highest row
  for (const n of [6, 8, 9, 10, 12, 13, 14, 15, 18, 20, 24, 28, 36, 48, 72]) add('wrapped', { cells: one('a\nb', size(n, WRAP)) })
  for (const n of [8, 14, 20, 36]) add('wrapped', { cells: one('Ag', size(n, WRAP)) })
  for (const n of [14, 24]) add('wrapped', { cells: one('a\nb\nc', size(n, WRAP)) })
  for (const n of [20, 25, 26]) add('wrapped', { cells: one(lines(n), WRAP) })
  // other fonts wrapped
  for (const font of [{ name: 'Arial', size: 10 }, { name: 'Arial', size: 16 }, { name: 'Times New Roman', size: 12 }, { name: 'Courier New', size: 10 }]) add('fonts', { cells: one('a\nb', { font, alignment: { wrap: true } }) })
  add('fonts', { cells: one('Ag', { font: { name: 'Arial', size: 10 }, alignment: { wrap: true } }) })
  add('fonts', { cells: one('a\nb\nc', { font: { name: 'Arial', size: 10 }, alignment: { wrap: true } }) })
  // the row's style smaller than the default: alone, under a cell of its own size, of the default's, wrapped
  add('rowstyle', { row: { look: size(8) } })
  add('rowstyle', { row: { look: size(8) }, cells: one('a') })
  add('rowstyle', { row: { look: size(8) }, cells: one('a', size(11)) })
  add('rowstyle', { row: { look: size(8, WRAP) }, cells: one('a\nb') })
  // a wrapped cell beside a larger font, and beside another wrapped one
  add('mixed', { cells: { 0: { value: 'a\nb', look: WRAP }, 1: { value: 'Ag', look: size(24) } } })
  add('mixed', { cells: { 0: { value: 'a\nb', look: WRAP }, 1: { value: 'Ag', look: size(28) } } })
  add('mixed', { cells: { 0: { value: 'a\nb', look: WRAP }, 1: { value: 'a\nb\nc', look: size(8, WRAP) } } })
  // where a line of 16 points breaks: a word of k letters b in a column of width 20 (120 points)
  for (let k = 10; k <= 16; k++) add('scale', { cells: one('b'.repeat(k), size(16, WRAP)), columns: { 0: { width: 20 } } })
  // the room at width 15 (95 points): a word and another of w points
  for (let w = 78; w <= 84; w++) add('room', { cells: one(`b ${textOf(w)}`, WRAP), columns: { 0: { width: 15 } } })
  // spaces at the break: after a line that fits, leading, two between; a hyphen inside a long word
  add('spaces', { cells: one(`b ${textOf(38)}   `, WRAP) })
  add('spaces', { cells: one(`${textOf(46)} b`, WRAP) })
  add('spaces', { cells: one(`   ${textOf(40)}`, WRAP) })
  add('spaces', { cells: one(`b  ${textOf(36)}`, WRAP) })
  add('spaces', { cells: one(`b  ${textOf(37)}`, WRAP) })
  add('hyphen', { cells: one('bbbbb-bbbbbbbbbb', WRAP) })
  add('hyphen', { cells: one('bbbbbbbbbbbbbbbbbbbbbbbb', WRAP) })
  // a boolean and an error wrapped in a narrow column
  add('value', { cells: one(true, WRAP), columns: { 0: { width: 2 } } })
  add('value', { cells: { 0: { formula: '1/0', look: WRAP } }, columns: { 0: { width: 2 } } })
  // a merged cell of a larger font, wrapped
  add('merged', { cells: one('a\nb', size(16, WRAP)), merged: [0] })
  // an empty hidden row, a row of height 0
  add('hidden', { row: { hidden: true } })
  add('hidden', { row: { height: 0 }, cells: one('a') })
  // bold text wrapped, by the widths of plain text?
  add('bold', { cells: one(SENTENCE, { font: { bold: true }, alignment: { wrap: true } }), columns: { 0: { width: 10 } } })
  add('bold', { cells: one(SENTENCE, { font: { bold: true }, alignment: { wrap: true } }), columns: { 0: { width: 25 } } })
  // long text over more widths
  for (const width of [7, 15, 25, 30, 50]) add('width', { cells: one(SENTENCE, WRAP), columns: { 0: { width } } })
  // through the object model: a formula giving a line break, one typed over two lines, text into a wrapped cell, a line
  // break into a larger font, into a merged cell, into a row with a taller cell, a carriage return, a short text into a
  // cell that wrapped three lines
  add('typed', { typed: [[0, '="a"&CHAR(10)&"b"']] })
  add('typed', { typed: [[0, '="a"&\n"b"']] })
  add('typed', { cells: { 0: { look: WRAP } }, typed: [[0, SENTENCE]] })
  add('typed', { cells: { 0: { look: size(16) } }, typed: [[0, 'a\nb']] })
  add('typed', { merged: [0], typed: [[0, 'a\nb']] })
  add('typed', { cells: { 0: { value: 'x', look: size(24) } }, typed: [[1, 'a\nb']] })
  add('typed', { typed: [[0, 'a\r\nb']], via: 'value' })
  add('typed', { cells: { 0: { value: 'a\nb\nc', look: WRAP } }, typed: [[0, 'x']] })
  add('typed', { typed: [[0, 'a\nb'], [0, 'x']] })
  return out
}
