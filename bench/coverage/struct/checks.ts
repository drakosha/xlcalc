/**
 * Held-out checks of rules that rest on few observations: moves the fitting set did not have (2-D strips,
 * moves onto another sheet's references, a middle block landing past a range's end), deleting and
 * inserting at the sheet's last row and column, copies with mixed `$` and with local names, a move of
 * empty cells over values, a copy that does not fit.
 */
import type { BookInput, CellInput, RangeRef } from '../../../src/contract.js'
import { rangeAt, type StructCase, type StructCases } from './cases.js'
import { windowBook } from './moves.js'
import { formulaCell as f } from '../stored.js'

/** E: values; F: references at the last row and column. */
function edgeBook(): BookInput {
  return {
    sheets: [
      { name: 'E', cells: { A1: { value: 1 } } },
      {
        name: 'F',
        cells: {
          A1: f('SUM(E!A1048576:A1048576)'),
          A2: f('SUM(E!A1048575:A1048576)'),
          A3: f('SUM(E!1048576:1048576)'),
          A4: f('SUM(E!A5:A1048576)'),
          A5: f('SUM(E!A4:A1048575)'),
          A6: f('SUM(E!XFD1:XFD1)'),
          A7: f('SUM(E!XFC1:XFD1)'),
          A8: f('SUM(E!XFD:XFD)'),
          A9: f('E!A1048575'),
          A10: f('SUM(E!B1048570:C1048576)'),
          A11: f('SUM(E!A1:B2)'),
        },
      },
    ],
  }
}

/** Copies with mixed `$` and a local name; a global name of the same name on another sheet. */
function copyBook(): BookInput {
  const s: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) s[`A${r}`] = { value: r }
  Object.assign(s, { C10: f('SUM(A$5:A6)'), D10: f('SUM($A5:A$9)'), E10: f('SUM(loc)+glob'), F10: f('SUM(A1:A2)+A$12') })
  return {
    sheets: [{ name: 'S', cells: s }, { name: 'T', cells: {} }, { name: 'V', cells: {} }],
    names: [
      { name: 'loc', ref: "'S'!$A$1:$A$3", scope: 'S' },
      { name: 'loc', ref: "'S'!$A$4" },
      { name: 'glob', ref: "'S'!$A$12" },
      { name: 'loc', ref: "'S'!$A$7", scope: 'V' },
    ],
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const move = (from: string, to: string, book = 'window') => {
    const t = rangeAt(to)
    cases.push({ id: `move-${from}-${to}`, book, op: { op: 'moveRange', from: rangeAt(from), to: { sheet: t.sheet, cell: t.range } } })
  }
  move('S!A5:A6', 'S!A12')
  move('S!A9:A11', 'S!A3')
  move('S!A2:A4', 'S!A3')
  move('S!A7:A9', 'S!A6')
  move('S!A1:A2', 'S!A11')
  move('S!A11:A12', 'S!A1')
  move('S!A4:A6', 'S!A1')
  move('S!A6:A8', 'S!A9')
  move('S!A3:B5', 'S!A9')
  move('S!A6:B6', 'S!C6')
  move('S!B2:B4', 'S!B8')
  move('U!A1:A3', 'S!A5')
  move('U!A1:B2', 'S!A11')
  const rows = (id: string, op: 'insertRows' | 'deleteRows', row: number, count = 1) => cases.push({ id, book: 'edge', op: { op, sheet: 'E', row, count } })
  const cols = (id: string, op: 'insertColumns' | 'deleteColumns', column: string, count = 1) => cases.push({ id, book: 'edge', op: { op, sheet: 'E', column, count } })
  rows('edge-delete-rows-1048575x2', 'deleteRows', 1048575, 2)
  rows('edge-delete-rows-5-to-end', 'deleteRows', 5, 1048572)
  rows('edge-insert-rows-1048576', 'insertRows', 1048576)
  rows('edge-delete-rows-1048574', 'deleteRows', 1048574)
  cols('edge-delete-cols-XFCx2', 'deleteColumns', 'XFC', 2)
  cols('edge-insert-cols-XFD', 'insertColumns', 'XFD')
  const copy = (id: string, from: string, to: string) => cases.push({ id, book: 'copy', op: { op: 'copyRange', from: rangeAt(from), to: rangeAt(to) } })
  copy('copy-mixed-up', 'S!C10:D10', 'S!C7')
  copy('copy-mixed-down', 'S!C10:D10', 'S!C13')
  copy('copy-local-name-T', 'S!E10', 'T!E10')
  copy('copy-local-name-V', 'S!E10', 'V!E10')
  cases.push({ id: 'move-local-name-T', book: 'copy', op: { op: 'moveRange', from: rangeAt('S!E10:F10'), to: { sheet: 'T', cell: 'A1' } } })
  cases.push({ id: 'move-empty-over-values', book: 'copy', op: { op: 'moveRange', from: rangeAt('S!H1:H2'), to: { sheet: 'S', cell: 'A1' } } })
  // last: Excel refuses a copy that does not fit (an error, if no window)
  copy('copy-off-sheet', 'S!C10:D10', 'S!XFD20')
  return { books: { window: windowBook({ A1: { value: 7 }, A2: { value: 8 }, A3: { value: 9 }, B1: { value: 1 } }), edge: edgeBook(), copy: copyBook() }, cases }
}
