/**
 * Text with line breaks entered on the keyboard of Excel for Mac (Control+Option+Return in the cell) and pasted: typed
 * into a cell and Return, into a row whose height was set by hand, a cell of a larger font, a merged cell; a formula
 * typed over two lines and one giving a line break; text typed into a wrapped cell; entered into a selection with
 * Control+Return, in a row of the default height and one set by hand; and text with line breaks pasted from the
 * clipboard, inside quotes and not, which Excel for Windows pastes too (`Worksheet.Paste`).
 */
import type { CellStyle } from '../../../src/contract.js'
import { collector, type RowCase } from './cases.js'

const WRAP: CellStyle = { alignment: { wrap: true } }
const SENTENCE = 'The quick brown fox jumps over the lazy dog. '.repeat(3).trim()

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  add('keys', { typed: [[0, 'abc']], via: 'keys' })
  add('keys', { typed: [[0, 'a\nb']], via: 'keys' })
  add('keys', { typed: [[0, 'a\nb\nc']], via: 'keys' })
  add('keys', { row: { height: 30 }, typed: [[0, 'a\nb\nc']], via: 'keys' })
  add('keys', { cells: { 0: { look: { font: { size: 16 } } } }, typed: [[0, 'a\nb']], via: 'keys' })
  add('keys', { merged: [0], typed: [[0, 'a\nb']], via: 'keys' })
  add('keys', { typed: [[0, '=1+\n2']], via: 'keys' })
  add('keys', { typed: [[0, '="a"&CHAR(10)&"b"']], via: 'keys' })
  add('keys', { cells: { 0: { look: WRAP } }, typed: [[0, SENTENCE]], via: 'keys' })
  add('keys', { cells: { 0: { value: 'a\nb\nc', look: WRAP } }, typed: [[0, 'x']], via: 'keys' })
  add('fill', { typed: [[0, 'a\nb']], via: 'fill', fillTo: 2 })
  add('fill', { row: { height: 30 }, typed: [[0, 'a\nb\nc']], via: 'fill', fillTo: 2 })
  add('fill', { typed: [[0, 'abc']], via: 'fill', fillTo: 2 })
  add('paste', { typed: [[0, '"a\nb"']], via: 'paste', below: true })
  add('paste', { typed: [[0, 'a\nb']], via: 'paste', below: true })
  add('paste', { typed: [[0, '"a\nb"\tc']], via: 'paste', fillTo: 1, below: true })
  add('paste', { typed: [[0, '"a b"']], via: 'paste' })
  add('paste', { typed: [[0, '"a\nb']], via: 'paste', below: true })
  add('paste', { typed: [[0, '"a""\nb"']], via: 'paste', below: true })
  return out
}
