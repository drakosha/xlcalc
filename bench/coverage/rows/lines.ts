/**
 * Wrapped lines, after `held` showed heights that are not a whole number of points a line: one to 26 lines of Calibri
 * 11 and more sizes, small sizes of few lines; half sizes not wrapped; a wrapped cell beside cells of larger and smaller
 * fonts, wrapped and not, as `held` showed rows higher than their highest cell; where a line of 20 and 24 points breaks.
 */
import type { CellStyle } from '../../../src/contract.js'
import { collector, type RowCase, type RowCell } from './cases.js'

const look = (size: number, wrap: boolean): CellStyle => ({ font: { size }, ...(wrap ? { alignment: { wrap: true } } : {}) })
const lines = (n: number) => Array.from({ length: n }, (_, i) => `l${i}`).join('\n')
const cell = (n: number, size: number, wrap = true): RowCell => ({ value: n === 1 && !wrap ? 'Ag' : lines(n), look: look(size, wrap) })

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  for (let n = 1; n <= 26; n++) add('l11', { cells: { 0: cell(n, 11) } })
  for (let n = 1; n <= 19; n++) add('l16', { cells: { 0: cell(n, 16) } })
  for (const n of [1, 2, 3, 4, 6, 8, 10, 12]) add('l24', { cells: { 0: cell(n, 24) } })
  for (const size of [6, 7, 8, 9, 10]) for (const n of [1, 2, 3, 4, 5, 8]) add(`l${size}`, { cells: { 0: cell(n, size) } })
  for (const size of [13.5, 15.5, 17.5, 18.5, 20.5]) add('half', { cells: { 0: cell(1, size, false) } })
  // a wrapped cell beside one of another font
  for (const n of [1, 2, 3]) for (const size of [14, 16, 20, 24, 36]) add('mixed', { cells: { 0: cell(n, 11), 1: cell(1, size, false) } })
  for (const size of [11, 24, 36]) add('mixed', { cells: { 0: cell(2, 16), 1: cell(1, size, false) } })
  add('mixed', { cells: { 0: cell(2, 11), 1: cell(3, 8) } })
  add('mixed', { cells: { 0: cell(5, 11), 1: cell(2, 24) } })
  add('mixed', { cells: { 0: cell(2, 11), 1: cell(1, 8, false) } })
  add('mixed', { cells: { 0: cell(3, 11), 1: cell(2, 16) } })
  // where a line of 20 and 24 points breaks in a column of width 20 (120 points)
  for (const k of [9, 10, 11, 12]) add('scale', { cells: { 0: { value: 'b'.repeat(k), look: look(20, true) } }, columns: { 0: { width: 20 } } })
  for (const k of [7, 8, 9, 10]) add('scale', { cells: { 0: { value: 'b'.repeat(k), look: look(24, true) } }, columns: { 0: { width: 20 } } })
  return out
}
