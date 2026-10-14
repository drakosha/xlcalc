/**
 * The widths Excel measures text by (Calibri 11, the default font of the books written here): a fill `*X`
 * after one digit repeats X across a wide column, so the count of X gives X's width against the column's;
 * a number too wide for its column shows `#` across it, so the count of `#` gives the room a column has;
 * whole numbers of 1 to 14 digits against columns from 1 to 12 characters give where digits stop fitting;
 * fills of `.` and of spaces in columns of fractional widths, and in the default one, give their room to a point.
 */
import { collector, sixths } from './cases.js'

/** The characters whose widths are measured by fill. */
export const GLYPHS = [
  ...'0123456789.,-+ $%()/:#E!^&~{}<>=?_\'',
  ...'abcdefghijklmnopqrstuvwxyz',
  ...'ABCDFGHIJKLMNOPQRSTUVWXYZ',
  ...'абвгдеёжзийклмнопрстуфхцчшщъыьэюя',
  ...'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ',
  ...'₽€£¥',
]

export function buildDisplayCases() {
  const { out, add } = collector()
  for (const g of GLYPHS) for (const width of [50, 100]) add('fill', 5, `0*${g}`, { width })
  for (const width of sixths(0.5, 25)) add('hashes', 1e20, '0', { width })
  for (const width of sixths(1, 12)) for (let k = 1; k <= 14; k++) add('digits', Number('8'.repeat(k)), '0', { width })
  const fractional = [undefined, 8.43, 8.7109375, 9.140625, 1.3, 2.29, 3.5, ...Array.from({ length: 24 }, (_, j) => 10 + j / 24)]
  for (const width of fractional) for (const g of ['.', ' ', '/', '+']) add('px', 5, `0*${g}`, width === undefined ? {} : { width })
  // 49 points of text in the default column, which has room for 48 or 49
  add('px', 5555555, '0"fr"')
  return out
}
