/**
 * The Format Painter double-clicked, which stays on over several targets until Escape, in both Excels: one cell over a
 * range and then two cells, a range onto two cells, a cell onto two cells one under the other (its conditional format
 * and validation in the source), a merged area twice, onto a merged area and then a cell, whole rows twice, onto another
 * sheet and back, three targets. The book of the set `painter`. The engine predicts a paste of formats onto each target
 * in turn (`painter-double.predicted.json`).
 */
import type { PasteCase, PasteCases } from './cases.js'
import { painterBook } from './painter.js'

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, more: string[], toSheet = 'S') =>
    cases.push({ id, book: 'painter', from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what: 'painter', more: more.map((m) => (m.includes('!') ? { sheet: m.split('!')[0], range: m.split('!')[1] } : { sheet: 'S', range: m })) })
  add('cell-to-range-and-cell', 'A1', 'F3:H5', ['F10:G11'])
  add('range-to-cells', 'A1:B2', 'F10', ['L1'])
  add('adjacent', 'A1', 'F3', ['F4'])
  add('merged-twice', 'A4:B4', 'F12', ['F16'])
  add('onto-merged-then-cell', 'A1', 'H14:I14', ['F18'])
  add('rows-twice', '6:6', '8:8', ['10:10'])
  add('other-sheet-and-back', 'A1:B2', 'B2', ['S!F20'], 'T')
  add('three', 'A1', 'F22', ['F24', 'F26'])
  return { books: { painter: painterBook() }, cases }
}
