/**
 * References across sheets and names of any formula under structure operations: rows and columns inserted or deleted
 * on a sheet the reference covers and on the formula's own, sheets renamed, deleted at an end or between, inserted
 * between the ends or past them, moved into, out of and across the span; copies and moves of formulas with such
 * references. Names hold formulas, unions, references without `$` and references across sheets.
 */
import type { BookInput, CellInput, StructureOp } from '../../../src/contract.js'
import { caseAdder, type StructCases } from './cases.js'
import { formulaCell as f } from '../stored.js'

function book(): BookInput {
  const data = (k: number): Record<string, CellInput> => ({ A1: { value: k }, A2: { value: k * 10 }, A3: { value: k * 100 }, B1: { value: k + 1 } })
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 }, C1: { value: 5 },
          F1: f('SUM(Jan:Mar!A1:A3)'),
          F2: f('SUM(Jan:Mar!A:A)'),
          F3: f('SUM(Feb:Mar!B1)'),
          F4: f('SUM(Jan:Mar!$A$2)'),
          F5: f('g'),
          F6: f('SUM(fm)'),
          F7: f('SUM(un)'),
          F8: f('rel'),
          F9: f('SUM(cb)'),
        },
      },
      { name: 'Jan', cells: data(1) },
      { name: 'Feb', cells: { ...data(2), D1: f('SUM(Jan:Mar!A1)') } },
      { name: 'Mar', cells: data(3) },
      { name: 'Z', cells: { A1: { value: 9 } } },
    ],
    names: [
      { name: 'g', ref: "SUM('Jan:Mar'!$A$1)" },
      { name: 'fm', ref: "'S'!$A$1:$A$3*2" },
      { name: 'un', ref: "'S'!$A$1:$A$2,'S'!$C$1" },
      { name: 'rel', ref: "'S'!XFD1" },
      { name: 'cb', ref: "'Jan:Mar'!$A$1:$A$2" },
    ],
  }
}

export function buildStructCases(): StructCases {
  const { cases, add } = caseAdder('cube')
  add('insert-rows-S', { op: 'insertRows', sheet: 'S', row: 2 })
  add('delete-rows-S', { op: 'deleteRows', sheet: 'S', row: 1 })
  add('insert-rows-Feb', { op: 'insertRows', sheet: 'Feb', row: 1 })
  add('delete-rows-Jan', { op: 'deleteRows', sheet: 'Jan', row: 2 })
  add('delete-columns-Mar', { op: 'deleteColumns', sheet: 'Mar', column: 'A' })
  add('rename-Jan', { op: 'renameSheet', sheet: 'Jan', name: 'January' })
  add('rename-Feb', { op: 'renameSheet', sheet: 'Feb', name: 'X' })
  add('delete-Jan', { op: 'deleteSheet', sheet: 'Jan' })
  add('delete-Feb', { op: 'deleteSheet', sheet: 'Feb' })
  add('delete-Mar', { op: 'deleteSheet', sheet: 'Mar' })
  add('insert-between', { op: 'insertSheet', name: 'New', before: 'Mar' })
  add('insert-before-first', { op: 'insertSheet', name: 'New', before: 'Jan' })
  add('insert-at-end', { op: 'insertSheet', name: 'New' })
  add('move-Z-in', { op: 'moveSheet', sheet: 'Z', before: 'Feb' })
  add('move-Feb-out', { op: 'moveSheet', sheet: 'Feb' })
  add('move-Jan-after-Mar', { op: 'moveSheet', sheet: 'Jan', before: 'Z' })
  add('move-Mar-first', { op: 'moveSheet', sheet: 'Mar', before: 'S' })
  add('copy-F1', { op: 'copyRange', from: { sheet: 'S', range: 'F1' }, to: { sheet: 'S', range: 'G2' } })
  add('move-F1', { op: 'moveRange', from: { sheet: 'S', range: 'F1' }, to: { sheet: 'S', cell: 'H5' } })
  add('move-A1-A2', { op: 'moveRange', from: { sheet: 'S', range: 'A1:A2' }, to: { sheet: 'S', cell: 'H1' } })
  return { books: { cube: book() }, cases }
}
