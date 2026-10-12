/**
 * The mark of a failed update on a link's data, held out. After an edit that entered formulas with a link Excel saved
 * that sheet's data with `refreshError="1"`, and cells the link holds no value of were `#REF!`; without the mark they
 * were blank (`edits-held`). Here: the mark on data with values, data without values with and without the mark, and
 * which data an edit marks (the sheet its formulas read, the link's other sheet, another link).
 *
 * `predicted` was committed before Excel ran the set. The guess: a cell the link holds no value of is blank on a sheet
 * whose data has no mark and `#REF!` on one whose data has it; an edit marks only the sheets its new formulas read.
 */
import type { LinksCase, LinksCases } from './cases.js'

const REF = { error: 'REF' } as const

const CASES: LinksCase[] = [
  {
    id: 'marked',
    about: 'data with the mark: a cell it holds, one it does not, a whole column',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A1', A2: '[1]Sheet1!A9', A3: 'SUM([1]Sheet1!A:A)', A4: 'ISBLANK([1]Sheet1!A9)' } },
    links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1, A2: 2, A3: 3 }, refreshError: true }] }],
    read: [{ sheet: 'S', range: 'A1:A4' }],
    predicted: { 'S!A1': 1, 'S!A2': REF, 'S!A3': REF, 'S!A4': false },
    alternatives: { 'S!A2': [0], 'S!A3': [6], 'S!A4': [true] },
  },
  {
    id: 'empty',
    about: 'data without values, with and without the mark, and a sheet without data',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Empty!A1', A2: '[1]Marked!A1', A3: '[1]None!A1', A4: 'SUM([1]Empty!A:A)' } },
    links: [
      {
        target: 'Book2.xlsx',
        sheets: [
          { name: 'Empty', cells: {} },
          { name: 'Marked', cells: {}, refreshError: true },
          { name: 'None', data: false },
        ],
      },
    ],
    read: [{ sheet: 'S', range: 'A1:A4' }],
    predicted: { 'S!A1': 0, 'S!A2': REF, 'S!A3': REF, 'S!A4': 0 },
    alternatives: { 'S!A1': [REF], 'S!A2': [0], 'S!A4': [REF] },
  },
  {
    id: 'scope',
    about: 'a fill down of a formula over Sheet1 of the first link: Sheet1, Sheet2 of the same link and the second link after it',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A9', A2: '[1]Sheet2!A9', A3: '[2]Data!A9', C1: '[1]Sheet1!A1' } },
    links: [
      {
        target: 'Book2.xlsx',
        sheets: [
          { name: 'Sheet1', cells: { A1: 1, A2: 2 } },
          { name: 'Sheet2', cells: { A1: 5 } },
        ],
      },
      { target: 'Data.xlsx', sheets: [{ name: 'Data', cells: { A1: 7 } }] },
    ],
    read: [
      { sheet: 'S', range: 'A1:A3' },
      { sheet: 'S', range: 'C1:C2' },
    ],
    ops: [{ op: 'fillDown', sheet: 'S', to: 'C1:C2' }],
    predicted: { 'S!A1': REF, 'S!A2': 0, 'S!A3': 0, 'S!C1': 1, 'S!C2': 2 },
    alternatives: { 'S!A1': [0], 'S!A2': [REF], 'S!A3': [REF] },
  },
]

export function buildLinksCases(): LinksCases {
  return { cases: CASES }
}
