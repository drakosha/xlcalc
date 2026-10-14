/**
 * Held out: SUBTOTAL and AGGREGATE over hidden rows in combinations the set `hidden` did not have, their answers
 * predicted by its rules and committed before Excel ran (`hidden-held.predicted.json`). Two-dimensional ranges over
 * hidden rows and columns of a filtered sheet, whole rows, nested formulas in visible and hidden rows, text,
 * booleans and errors in hidden rows, the functions with k over hidden rows, a filter applied after opening.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { CellInput, SheetInput } from '../../../src/contract.js'
import type { WorkbookCase } from './cases.js'
import { formulaCell as f } from '../stored.js'

/** S: A1:C1 a header, A2:C13 values (A powers of two, B text and booleans, C numbers), D2:D13 the filter's keys; a nested SUBTOTAL in A9 and A12. */
function data(): SheetInput {
  const cells: Record<string, CellInput> = { A1: { value: 'a' }, B1: { value: 'b' }, C1: { value: 'c' }, D1: { value: 'k' } }
  const keys = 'xyxxyxxyxyxx'
  for (let i = 0; i < 12; i++) {
    const r = i + 2
    cells[`A${r}`] = { value: 2 ** i }
    cells[`B${r}`] = { value: i % 3 === 0 ? 't' : i % 3 === 1 ? true : 2 ** (i + 12) }
    cells[`C${r}`] = { value: (i % 5) + 0.5 }
    cells[`D${r}`] = { value: keys[i] }
  }
  // nested formulas: one in a visible row, one in a hidden row; an error in a hidden row
  Object.assign(cells, { A9: f('SUBTOTAL(9,A2:A3)'), A12: f('AGGREGATE(9,4,A2:A3)*2'), C6: f('1/0') })
  return { name: 'S', cells }
}

const FS = [
  'SUBTOTAL(9,S!A2:C13)',
  'SUBTOTAL(109,S!A2:C13)',
  'SUBTOTAL(2,S!A2:C13)',
  'SUBTOTAL(102,S!A2:C13)',
  'SUBTOTAL(3,S!A2:C13)',
  'SUBTOTAL(103,S!A2:C13)',
  'SUBTOTAL(1,S!A2:A13)',
  'SUBTOTAL(101,S!A2:A13)',
  'SUBTOTAL(4,S!A2:B13)',
  'SUBTOTAL(104,S!A2:B13)',
  'SUBTOTAL(5,S!C2:C13)',
  'SUBTOTAL(105,S!C2:C13)',
  'SUBTOTAL(6,S!C2:C5)',
  'SUBTOTAL(106,S!C2:C5)',
  'SUBTOTAL(7,S!A2:A13)',
  'SUBTOTAL(107,S!A2:A13)',
  'SUBTOTAL(10,S!A2:A13)',
  'SUBTOTAL(110,S!A2:A13)',
  'SUBTOTAL(9,S!3:8)',
  'SUBTOTAL(109,S!3:8)',
  'SUBTOTAL(109,S!A2:A5,S!A10:A13,S!B2:B13)',
  'SUBTOTAL(9,S!C2:C13)',
  'SUBTOTAL(109,S!C2:C13)',
  'AGGREGATE(9,0,S!C2:C13)',
  'AGGREGATE(9,1,S!C2:C13)',
  'AGGREGATE(9,2,S!C2:C13)',
  'AGGREGATE(9,3,S!C2:C13)',
  'AGGREGATE(9,6,S!C2:C13)',
  'AGGREGATE(9,7,S!C2:C13)',
  'AGGREGATE(3,5,S!A2:C13)',
  'AGGREGATE(2,1,S!A2:C13)',
  'AGGREGATE(14,1,S!A2:A13,1)',
  'AGGREGATE(14,5,S!A2:A13,2)',
  'AGGREGATE(15,3,S!A2:A13,1)',
  'AGGREGATE(16,5,S!A2:A13,0.5)',
  'AGGREGATE(17,7,S!A2:A13,3)',
  'AGGREGATE(12,5,S!A2:A13)',
  'AGGREGATE(13,1,S!C2:C5)',
  'AGGREGATE(15,6,S!A2:A13/(S!D2:D13="x"),1)',
  'AGGREGATE(14,7,S!A2:A13/(S!D2:D13="x"),2)',
]

const book = { sheets: [data(), { name: 'F', cells: Object.fromEntries(FS.map((x, i) => [`A${i + 1}`, f(x)])) }] }
const read = [{ sheet: 'F', range: `A1:A${FS.length}` }]

const PREDICTED = JSON.parse(readFileSync(new URL('./hidden-held.predicted.json', import.meta.url), 'utf8')) as Record<string, Record<string, Val>[]>

export function buildWorkbookCases(): WorkbookCase[] {
  const cases: WorkbookCase[] = [
    { id: 'manual', about: 'rows 3, 6:7 and 12 hidden by hand, column B hidden', book, read, file: { sheets: { S: { hiddenRows: ['3', '6:7', '12'], hiddenColumns: ['B'] } } } },
    { id: 'filtered', about: 'a filter keeping x hides rows 3, 6, 9 and 11; row 13 hidden too', book, read, file: { sheets: { S: { filter: { range: 'A1:D13', keep: { 3: ['x'] } }, hiddenRows: ['3', '6', '9', '11', '13'] } } } },
    {
      id: 'filter-later',
      about: 'rows 6:7 hidden by hand, then Excel filters to x, then row 6 shown',
      book,
      read,
      file: { sheets: { S: { hiddenRows: ['6:7'] } } },
      steps: ['full', 'read', { filter: { sheet: 'S', range: 'A1:D13', field: 4, criterion: 'x' } }, 'read', { hide: { sheet: 'S', rows: '6', hidden: false } }, 'read'],
    },
  ]
  return cases.map((c) => (PREDICTED[c.id] ? { ...c, predicted: PREDICTED[c.id] } : c))
}
