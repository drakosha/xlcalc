/**
 * Held out after `fills-held`: fills whose colour is there but names none, Excel's automatic one (`auto="1"`) or the
 * system's (`indexed="64"`, `65`), where `readFill` reads a cell's solid fill as white and a dxf's as no fill, and a
 * dxf's pattern fill with nothing in it. The book and its placeholders are `fills-held`'s: A1:A3 in cell formats of
 * their own, B1:C4 under one rule per row, C in yellow. Predicted before Excel ran it (`fills-held-2.predicted.json`).
 */
import type { RawEdit } from '../xlsx/sent.js'
import type { CondCases } from './cases.js'
import { book, cellFill, dxfFill, STYLES } from './fills-held.js'

const RED = 'FFFF0000'

/** The fills of A1:A3 as the file holds them. */
const CELL_FILLS = [
  `<patternFill patternType="solid"><fgColor auto="1"/><bgColor rgb="${RED}"/></patternFill>`,
  '<patternFill patternType="solid"><fgColor indexed="64"/><bgColor indexed="64"/></patternFill>',
  `<patternFill patternType="solid"><fgColor indexed="65"/><bgColor rgb="${RED}"/></patternFill>`,
]

/** The fills of the rules over B1:C4, one per row. */
const DXF_FILLS = [
  '<patternFill patternType="solid"><fgColor indexed="2"/><bgColor auto="1"/></patternFill>',
  '<patternFill><bgColor indexed="64"/></patternFill>',
  '<patternFill/>',
  '<patternFill patternType="none"><bgColor auto="1"/></patternFill>',
]

export function buildCondCases(): CondCases {
  const raw: RawEdit[] = [...CELL_FILLS.map((to, i) => ({ part: STYLES, from: cellFill(i), to })), ...DXF_FILLS.map((to, i) => ({ part: STYLES, from: dxfFill(CELL_FILLS.length + i), to }))]
  return { books: { fills: book(CELL_FILLS.length, DXF_FILLS.length) }, cases: [{ id: 'fills', book: 'fills', read: [['S', 'A1:C4']], raw }] }
}
