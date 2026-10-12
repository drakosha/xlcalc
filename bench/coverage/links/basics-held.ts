/**
 * Links to other workbooks, held out: what a formula reads of another book's values stored in the file. One cell, a
 * range in SUM, a whole column and a whole row, a cell the link holds no value of, text, booleans and errors stored,
 * sheets with a space and in Cyrillic, the other book's names (of the book, of a sheet, a constant, a lost one), this
 * book's names over a link, INDEX and MATCH, VLOOKUP, XLOOKUP, SUMIFS, COUNTIF and SUMPRODUCT over a linked range,
 * INDIRECT and OFFSET into a linked book, a dynamic array over a link, a fill down and a copy of formulas with links, a
 * sheet of this book with the other book's sheet's name renamed, rows inserted above a formula with a link; data Excel
 * marks as a failed update and a sheet without data; two links in one book; the forms of a link's target.
 *
 * `predicted` was committed before Excel ran the set: the rule over the stored values (another book's cell is its
 * stored value, a cell it holds no value of is blank, a link's range is the stored cells in their places), and where
 * the rule was not known the answers thought possible in `alternatives`.
 */
import type { CachedValue, LinkedBook, LinksCase, LinksCases } from './cases.js'

const BOOK2: LinkedBook = {
  target: 'Book2.xlsx',
  sheets: [
    {
      name: 'Sheet1',
      cells: { A1: 1, A2: 2, A3: 3, B1: 'abc', B2: true, B3: { error: 'NA' }, B4: { error: 'DIV0' }, C1: 0.2, C2: 7, D1: 'x', D2: 'y', D3: 'x', E1: 10, F1: 20 },
    },
    { name: 'My Sheet', cells: { A1: 11 } },
    { name: 'Лист 1', cells: { A1: 12 } },
  ],
  names: [
    { name: 'Rate', ref: '=Sheet1!$C$1' },
    { name: 'Local', ref: '=Sheet1!$C$2', sheet: 0 },
    { name: 'Konst', ref: '=5' },
    { name: 'Bad', ref: '=#REF!' },
  ],
}

/** The formulas of S!A1 down with their predicted values. */
const CACHE: [formula: string, predicted: CachedValue | null, alternatives?: (CachedValue | null)[]][] = [
  ['[1]Sheet1!A1+1', 2],
  ['SUM([1]Sheet1!A1:A3)', 6],
  ['SUM([1]Sheet1!A:A)', 6],
  ['[1]Sheet1!A9', 0, [{ error: 'REF' }, { error: 'NA' }]],
  ['[1]Sheet1!A9&"|"', '|', ['0|', { error: 'REF' }]],
  ['ISBLANK([1]Sheet1!A9)', true, [false]],
  ['[1]Sheet1!B1', 'abc'],
  ['[1]Sheet1!B2', true],
  ['[1]Sheet1!B3', { error: 'NA' }],
  ['IFERROR([1]Sheet1!B4,"e")', 'e'],
  ["'[1]My Sheet'!A1", 11],
  ["'[1]Лист 1'!A1", 12],
  ['[1]!Rate', 0.2],
  ['[1]Sheet1!Local', 7],
  ['[1]!Konst', 5, [{ error: 'REF' }, { error: 'NAME' }]],
  ['Ext*2', 4],
  ['INDEX([1]Sheet1!A1:A3,MATCH(3,[1]Sheet1!A1:A3,0))', 3],
  ['VLOOKUP(2,[1]Sheet1!A1:B3,2,FALSE)', true],
  ['SUMIFS([1]Sheet1!A1:A3,[1]Sheet1!D1:D3,"x")', { error: 'VALUE' }, [4]],
  ['COUNTIF([1]Sheet1!D1:D3,"x")', { error: 'VALUE' }, [2]],
  ['SUMPRODUCT(([1]Sheet1!D1:D3="x")*[1]Sheet1!A1:A3)', 4],
  ['INDIRECT("[1]Sheet1!A1")', { error: 'REF' }],
  ['OFFSET([1]Sheet1!A1,1,0)', { error: 'VALUE' }, [2, { error: 'REF' }]],
  ['SUM([1]Sheet1!1:1)', 31.2],
  ['ROWS([1]Sheet1!A1:A3)', 3],
  ['ROW([1]Sheet1!B5)', 5],
  ['ISREF([1]Sheet1!A1)', true],
  ['COUNTA([1]Sheet1!A1:B4)', 7],
  ['COUNTBLANK([1]Sheet1!A1:A5)', { error: 'VALUE' }, [2]],
  ['[1]!Bad', { error: 'REF' }],
  ['SUM(ExtR)', 6],
  ['MATCH(20,[1]Sheet1!1:1,0)', 6],
  ['SUM([1]Sheet1!A1:A3,Sheet1!A1)', 1006],
  ['_xlfn.XLOOKUP(2,[1]Sheet1!A:A,[1]Sheet1!B:B)', true],
  ['AVERAGE([1]Sheet1!A:A)', 2],
  ['SUMIF([1]Sheet1!D1:D3,"x",[1]Sheet1!A1:A3)', { error: 'VALUE' }, [4]],
  ['INDIRECT("\'[Book2.xlsx]Sheet1\'!A1")', { error: 'REF' }],
]

const cacheCells = Object.fromEntries(CACHE.map(([f], i) => [`A${i + 1}`, f]))
const cachePredicted = Object.fromEntries(CACHE.map(([, v], i) => [`S!A${i + 1}`, v]))
const cacheAlternatives = Object.fromEntries(CACHE.flatMap(([, , alt], i) => (alt ? [[`S!A${i + 1}`, alt]] : [])))
const ARRAY_AT = CACHE.length + 1

const CASES: LinksCase[] = [
  {
    id: 'cache',
    about: 'formulas over one link: cells, ranges, whole columns and rows, names, lookups, criteria, INDIRECT, OFFSET, a dynamic array; a fill down and a copy, a sheet of this book named like the link’s renamed',
    sheets: [
      { name: 'S', cells: {} },
      { name: 'Sheet1', cells: { A1: { value: 1000 } } },
    ],
    linked: {
      S: {
        ...cacheCells,
        [`A${ARRAY_AT}`]: { formula: '[1]Sheet1!A1:A3*2', array: `A${ARRAY_AT}:A${ARRAY_AT + 2}` },
        C1: '[1]Sheet1!A1',
        C5: '[1]Sheet1!$A$1',
      },
    },
    names: [
      { name: 'Ext', ref: '[1]Sheet1!$A$2' },
      { name: 'ExtR', ref: '[1]Sheet1!$A$1:$A$3' },
    ],
    links: [BOOK2],
    read: [
      { sheet: 'S', range: `A1:A${ARRAY_AT + 2}` },
      { sheet: 'S', range: 'C1:D10' },
    ],
    ops: [
      { op: 'fillDown', sheet: 'S', to: 'C1:C3' },
      { op: 'fillDown', sheet: 'S', to: 'C5:C6' },
      { op: 'copyRange', sheet: 'S', range: 'C1', to: 'D10' },
      { op: 'renameSheet', sheet: 'Sheet1', to: 'Local1' },
    ],
    predicted: {
      ...cachePredicted,
      [`S!A${ARRAY_AT}`]: 2,
      [`S!A${ARRAY_AT + 1}`]: 4,
      [`S!A${ARRAY_AT + 2}`]: 6,
      'S!C1': 1,
      'S!C2': 2,
      'S!C3': 3,
      'S!C5': 1,
      'S!C6': 1,
      'S!D10': 0,
    },
    alternatives: { ...cacheAlternatives, 'S!D10': [{ error: 'REF' }] },
  },
  {
    id: 'refresh-error',
    about: 'data Excel marks as a failed update, and a sheet the link names without data',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A1+1', A2: '[1]Other!A1', A3: 'ISBLANK([1]Other!A1)' } },
    links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5 }, refreshError: true }, { name: 'Other', data: false }] }],
    read: [{ sheet: 'S', range: 'A1:A3' }],
    predicted: { 'S!A1': 6, 'S!A2': 0, 'S!A3': true },
    alternatives: { 'S!A1': [{ error: 'REF' }], 'S!A2': [{ error: 'REF' }], 'S!A3': [false] },
  },
  {
    id: 'two-links',
    about: 'two links in one book',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A1+[2]Data!A1', A2: '[2]Data!A1' } },
    links: [
      { target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5 } }] },
      { target: 'Data.xlsx', sheets: [{ name: 'Data', cells: { A1: 7 } }] },
    ],
    read: [{ sheet: 'S', range: 'A1:A2' }],
    predicted: { 'S!A1': 12, 'S!A2': 7 },
  },
  {
    id: 'inserted',
    about: 'rows inserted above a formula with a link: the formula moves, its link does not',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A3: '[1]Sheet1!A3', B3: 'A3' } },
    links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A3: 3, A4: 4 } }] }],
    read: [{ sheet: 'S', range: 'A1:B5' }],
    ops: [{ op: 'insertRows', sheet: 'S', range: 'A1' }],
    predicted: { 'S!A4': 3, 'S!B4': 3 },
  },
  {
    id: 'targets',
    about: 'the forms of a link’s target: a file name, a relative path, file:///, an absolute path, a web address, a folder and a name with a space and Cyrillic',
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: Object.fromEntries([1, 2, 3, 4, 5, 6].map((k) => [`A${k}`, `[${k}]Sheet1!A1`])) },
    links: ['Book2.xlsx', '../data/Book3.xlsx', 'file:///C:\\data\\Book4.xlsx', '/Users/x/data/Book5.xlsx', 'https://example.com/share/Book6.xlsx', 'Sub Folder/Книга 7.xlsx'].map((target, i) => ({
      target,
      sheets: [{ name: 'Sheet1', cells: { A1: i + 1 } }],
    })),
    read: [{ sheet: 'S', range: 'A1:A6' }],
    predicted: Object.fromEntries([1, 2, 3, 4, 5, 6].map((k) => [`S!A${k}`, k])),
  },
]

export function buildLinksCases(): LinksCases {
  return { cases: CASES }
}
