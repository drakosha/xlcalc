/**
 * The two forms of a formula over a link, each in a file made by hand: the old kind (`<f>`) and a dynamic array (`cm`,
 * `t="array"`), the same formulas in the same order, and those Excel saves when they are typed. Whether the old kind, which
 * Excel before 365 reads and edits as it is, can be written for a formula over a link.
 *
 * `predicted` was committed before Excel ran the set. The guess: where each reference into a link is one cell, the two forms
 * compute alike, what the engine computes, with no `@` shown and no repair window, a function that takes a reference too
 * (`COUNTBLANK`, `SUMIF` and `COUNTIF` #VALUE!, as over a closed workbook);
 * ranges (`ranges-*`) differ as they did (`SUM([1]Sheet1!A:A)`, a range alone), which is why these are files of their own:
 * a refused file costs only them.
 */
import type { CachedValue, LinkedBook, LinkedFormula, LinksCase, LinksCases } from './cases.js'

const NA = { error: 'NA' } as const
const VALUE = { error: 'VALUE' } as const
const REF = { error: 'REF' } as const

const LINK: LinkedBook = {
  target: 'Book2.xlsx',
  sheets: [
    { name: 'Sheet1', cells: { A1: 5, A2: 'txt', A3: true, A4: NA, A5: null, A6: '', B1: 7, B2: 8, B3: 9 } },
    { name: 'Marked', refreshError: true, cells: { A1: 3 } },
    { name: 'NoData', data: false },
  ],
}

/** One cell in each reference into the link, and what is predicted of it in both forms. */
const CELLS: [string, CachedValue | null][] = [
  ['[1]Sheet1!A1', 5],
  ['[1]Sheet1!A2', 'txt'],
  ['[1]Sheet1!A3', true],
  ['[1]Sheet1!A4', NA],
  ['[1]Sheet1!A5', 0],
  ['[1]Sheet1!A6', ''],
  ['[1]Sheet1!A9', 0],
  ['[1]Sheet1!$B$1', 7],
  ['[1]Sheet1!A1*2+[1]Sheet1!B1', 17],
  ['[1]Sheet1!A2&"|"', 'txt|'],
  ['[1]Sheet1!A9&"|"', '|'],
  ['-[1]Sheet1!A1', -5],
  ['[1]Sheet1!A1=5', true],
  ['[1]Sheet1!A1+C1', 7],
  ['SUM([1]Sheet1!A1)', 5],
  ['SUM([1]Sheet1!A1,[1]Sheet1!B1,1)', 13],
  ['MAX([1]Sheet1!A1,2)', 5],
  ['IF([1]Sheet1!A1>0,[1]Sheet1!B1,0)', 7],
  ['IFERROR([1]Sheet1!A4,0)', 0],
  ['ISBLANK([1]Sheet1!A9)', true],
  ['ISBLANK([1]Sheet1!A5)', true],
  ['ISNUMBER([1]Sheet1!A1)', true],
  ['LEN([1]Sheet1!A2)', 3],
  ['TEXT([1]Sheet1!A1,"0.00")', '5.00'],
  ['COUNTBLANK([1]Sheet1!A9)', VALUE],
  ['COUNTA([1]Sheet1!A1)', 1],
  ['COUNT([1]Sheet1!A9)', 0],
  ['ROW([1]Sheet1!A5)', 5],
  ['COLUMN([1]Sheet1!B1)', 2],
  ['ROWS([1]Sheet1!A1)', 1],
  ['OFFSET([1]Sheet1!A1,1,0)', VALUE],
  ['INDEX([1]Sheet1!A1,1)', 5],
  ['VLOOKUP(5,[1]Sheet1!A1,1,FALSE)', 5],
  ['MATCH(5,[1]Sheet1!A1,0)', 1],
  ['SUMIF([1]Sheet1!A1,">0")', VALUE],
  ['COUNTIF([1]Sheet1!A1,5)', VALUE],
  ['[1]Marked!A1', 3],
  ['[1]Marked!A9', REF],
  ['SUM([1]Marked!A9)', REF],
  ['[1]NoData!A1', REF],
  ['SUM([1]NoData!A1)', REF],
]

/** Ranges, each with the cells under it free for a dynamic array to spill into. */
const RANGES: [string, string][] = [
  ['[1]Sheet1!A1:A3', 'A1:A3'],
  ['[1]Sheet1!A1:A1', 'A6'],
  ['SUM([1]Sheet1!A1:A1)', 'A11'],
  ['SUM([1]Sheet1!A1:A3)', 'A16'],
  ['SUM([1]Sheet1!A:A)', 'A21'],
  ['COUNTBLANK([1]Sheet1!A1:A5)', 'A26'],
  ['INDEX([1]Sheet1!B1:B3,2)', 'A31'],
  ['VLOOKUP(5,[1]Sheet1!A1:B3,2,FALSE)', 'A36'],
]

const own = (cells: Record<string, unknown>): LinksCase['sheets'] => [{ name: 'S', cells: { C1: { value: 2 }, ...cells } as never }]
const first = (range: string) => range.split(':')[0]

const cellCases = (id: string, about: string, form: (f: string, at: string) => LinkedFormula): LinksCase => ({
  id,
  about,
  sheets: own({}),
  linked: { S: Object.fromEntries(CELLS.map(([f], i) => [`A${i + 1}`, form(f, `A${i + 1}`)])) },
  links: [LINK],
  read: [{ sheet: 'S', range: `A1:A${CELLS.length}` }],
  predicted: Object.fromEntries(CELLS.map(([, v], i) => [`S!A${i + 1}`, v])),
})

const rangeCases = (id: string, about: string, form: (f: string, at: string) => LinkedFormula): LinksCase => ({
  id,
  about,
  sheets: own({}),
  linked: { S: Object.fromEntries(RANGES.map(([f, at]) => [first(at), form(f, at)])) },
  links: [LINK],
  read: [{ sheet: 'S', range: 'A1:A40' }],
})

const plain = (f: string): LinkedFormula => f
const dynamic = (f: string, at: string): LinkedFormula => ({ formula: f, array: at })

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      cellCases('cells-plain', 'each reference one cell, of the old kind', plain),
      cellCases('cells-dynamic', 'each reference one cell, as a dynamic array', dynamic),
      rangeCases('ranges-plain', 'ranges, of the old kind', plain),
      rangeCases('ranges-dynamic', 'ranges, as a dynamic array', dynamic),
    ],
  }
}
