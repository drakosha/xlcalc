/**
 * AutoFit of a column: text of every width from 3 to 60 points and some wider (where the width comes from and its
 * rounding, up to the largest column), numbers in General and by formats (dates, money, padding `_)`, fills, text
 * sections), booleans and errors, several cells, empty columns, a column narrowed, merged cells, a hidden row, fonts,
 * formula results and a spill, a hidden column, a column's own format.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, HIDDEN, type FitCase } from './cases.js'

/** Text of `w` points of Calibri 11 on the Mac (`docs/formats.md`, Width): `b` is 6, the last letter makes up the rest. */
export function textOf(w: number): string {
  const tail: Record<number, string> = { 0: '', 1: 'D', 2: 'w', 3: 'i', 4: 'r', 5: 'a' }
  const q = Math.floor(w / 6)
  const rem = w % 6
  // 7 and 8 points take the place of a 6
  return rem === 1 || rem === 2 ? 'b'.repeat(q - 1) + tail[rem] : 'b'.repeat(q) + tail[rem]
}

const one = (v: CellInput['value'], format?: CellInput['format']): Record<number, CellInput> => ({ 1: { value: v, ...(format !== undefined ? { format } : {}) } })
const formula = (f: string, format?: CellInput['format']): Record<number, CellInput> => ({ 1: { formula: f, ...(format !== undefined ? { format } : {}) } })

const GENERAL = [1, 12, 123456, 12345678, 1234567890, 12345678901, 123456789012, -5, -123456789, 0.5, 3.14159265358979, 1e-10, 1e100, 0.000123456789, 123.456, -0.001, 99999999999]

const FORMATS: [number, string | number][] = [
  [1234567.891, '#,##0.00'],
  [0.5, '0%'],
  [0.123456, '0.00%'],
  [1234.5, '$#,##0.00_);($#,##0.00)'],
  [-1234.5, '$#,##0.00_);($#,##0.00)'],
  [1234567, '#,##0_);[Red](#,##0)'],
  [-1234567, '#,##0_);[Red](#,##0)'],
  [123456, '0.00E+00'],
  [1.25, '# ?/?'],
  [46096, 14],
  [46096, 15],
  [46096.75, 22],
  [1234.5, 5],
  [1234.5, 44],
  [-1234.5, 44],
  [46096, 'd mmmm yyyy'],
  [46096, 'dddd, mmmm d, yyyy'],
  [46096, '[$-F800]dddd, mmmm dd, yyyy'],
  [0.75, 'h:mm:ss AM/PM'],
  [2.5, '[h]:mm'],
  [46096.5, 'yyyy-mm-dd hh:mm:ss'],
  [5, '0*-'],
  [5, '"Total: "0*.'],
  [12345, '@'],
  [5, '"Total: "0'],
  [5, ';;;'],
  [1234.5, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  [0, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  [-1, 14],
  [1234567, '0_)'],
  [1234567, '0_W'],
]

export function buildFitCases(): FitCase[] {
  const { out, add } = collector()
  // text by its width in points: where the width comes from, then wide text up to the largest column (255 digits)
  for (let w = 3; w <= 60; w++) add('text', { cells: one(textOf(w)) })
  for (const w of [66, 72, 97, 121, 150, 199, 300, 600, 1000, 1524, 1525, 1526, 1530, 1536, 2000]) add('text', { cells: one(textOf(w)) })
  add('glyphs', { cells: one('Mixed WIDTHS: 1/2, (x) & 100%!') })
  add('glyphs', { cells: one('Цена товара: 1 234,56 ₽') })
  add('glyphs', { cells: one('  spaces  ') })
  add('glyphs', { cells: one('中文字符') })
  for (const n of GENERAL) add('general', { cells: one(n) })
  for (const [v, f] of FORMATS) add('format', { cells: one(v, f) })
  add('word', { cells: one(true) })
  add('word', { cells: one(false) })
  for (const f of ['1/0', 'NA()', '"a"+1', 'SQRT(-1)', 'FOO(1)']) add('word', { cells: formula(f) })
  add('word', { cells: one('TRUE', '@') })
  // several cells: the widest counts, whatever its row and kind
  add('many', { cells: { 1: { value: 'a' }, 2: { value: 'a longer text here' }, 3: { value: 12345 } } })
  add('many', { cells: { 1: { value: 46096, format: 'd mmmm yyyy' }, 2: { value: 'short' }, 40: { value: 1234567.5, format: '#,##0.00' } } })
  add('many', { cells: { 1: { value: 'abc' }, 5: { formula: '1/0' }, 9: { value: true } } })
  // nothing to show: an empty column with a width set by hand, one without, formatted blanks, empty text
  add('empty', { column: { width: 20 } })
  add('empty', { cells: { 1: { format: '0.00' } } })
  add('empty', { cells: { 1: { format: '0.00' } }, column: { width: 3 } })
  add('empty', { cells: formula('""') })
  add('empty', { cells: one(''), column: { width: 20 } })
  add('empty', { cells: one(-1, 14), column: { width: 20 } })
  // narrower and wider than the column set by hand
  add('narrow', { cells: one('abc'), column: { width: 30 } })
  add('narrow', { cells: one('abcdefghijklmnopqrstuvwxyz'), column: { width: 2 } })
  add('narrow', { cells: one(46096, 14), column: { width: 3 } })
  // merged cells: across into the next column, down into the next rows
  add('merged', { cells: { 1: { value: 'a long text in a merged cell' }, 2: { value: 'ab' } }, merged: [1] })
  add('merged', { cells: { 1: { value: 'a long text in a merged cell' }, 3: { value: 'ab' } }, mergedDown: [1, 2] })
  add('merged', { cells: { 1: { value: 'a long text in a merged cell' } }, merged: [1] })
  // a hidden row
  add('hidden', { cells: { 1: { value: 'ab' }, [HIDDEN]: { value: 'a long text in a hidden row' } } })
  add('hidden', { cells: { [HIDDEN]: { value: 'a long text in a hidden row' } } })
  add('hidden', { cells: one('a long text in a hidden column'), column: { hidden: true } })
  add('hidden', { cells: one('a long text in a hidden column'), column: { hidden: true, width: 5 } })
  // fonts other than the default: bold, size 16
  add('font', { cells: { 1: { value: 'Bold text of some width', style: 0 } } })
  add('font', { cells: { 1: { value: 'Larger text', style: 1 } } })
  add('font', { cells: { 1: { value: 1234567.891, format: '#,##0.00', style: 0 } } })
  // formula results, a spill, the column's own format
  add('formula', { cells: formula('REPT("a",20)') })
  add('formula', { cells: formula('DATE(2026,3,15)', 14) })
  add('formula', { cells: formula('SEQUENCE(3,1,1000000,1000000)') })
  add('formula', { cells: { 1: { formula: 'SEQUENCE(3,1,1000000,1000000)' }, 2: { format: '#,##0.00' } } })
  add('colformat', { cells: one(1234567.891), column: { format: '#,##0.00' } })
  add('colformat', { cells: one(46096), column: { format: 'dddd, mmmm d, yyyy' } })
  return out
}
