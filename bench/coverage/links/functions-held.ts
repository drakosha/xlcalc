/**
 * Functions over another workbook's stored values, held out. `basics-held` showed SUMIFS, SUMIF, COUNTIF, COUNTBLANK and
 * OFFSET giving `#VALUE!` over a linked range, while SUM, INDEX, MATCH, VLOOKUP, XLOOKUP and SUMPRODUCT compute. Here the
 * other functions whose argument takes only a reference (the criteria functions' ranges, RANK's list, SUBTOTAL's and
 * AGGREGATE's references, a database function's range, CELL, AREAS), a linked value as a criterion, a linked range
 * through a name of this book, and more functions that read values.
 *
 * `predicted` was committed before Excel ran the set. The guess: an argument that takes only a reference (the
 * engine's table `NOT_OMITTED`, COUNTBLANK's, OFFSET's, AGGREGATE's and a database function's range) is `#VALUE!` when it
 * is a linked range; any other argument reads the stored values.
 */
import type { CachedValue, LinksCase, LinksCases } from './cases.js'

const VALUE = { error: 'VALUE' } as const

/** The formulas of S!A1 down, their predicted values and the other answers thought possible. */
const FORMULAS: [formula: string, predicted: CachedValue | null, alternatives?: (CachedValue | null)[]][] = [
  ['COUNTIFS([1]Sheet1!B2:B4,"x")', VALUE, [2]],
  ['AVERAGEIF([1]Sheet1!B2:B4,"x",[1]Sheet1!C2:C4)', VALUE, [20]],
  ['AVERAGEIFS([1]Sheet1!C2:C4,[1]Sheet1!B2:B4,"x")', VALUE, [20]],
  ['_xlfn.MAXIFS([1]Sheet1!C2:C4,[1]Sheet1!B2:B4,"x")', VALUE, [30]],
  ['_xlfn.MINIFS([1]Sheet1!C2:C4,[1]Sheet1!B2:B4,"x")', VALUE, [10]],
  ['RANK(2,[1]Sheet1!A1:A3)', VALUE, [2]],
  ['_xlfn.RANK.EQ(2,[1]Sheet1!A1:A3)', VALUE, [2]],
  ['SUBTOTAL(9,[1]Sheet1!A1:A3)', VALUE, [6]],
  ['_xlfn.AGGREGATE(9,4,[1]Sheet1!A1:A3)', VALUE, [6]],
  ['DSUM([1]Sheet1!B1:C4,"v",F1:F2)', VALUE, [40]],
  ['ISTEXT(CELL("address",[1]Sheet1!A2))', true, [false]],
  ['AREAS([1]Sheet1!A1:A3)', 1, [VALUE]],
  ['INDEX([1]Sheet1!A1:A3,2)', 2],
  ['SUM(INDEX([1]Sheet1!A1:A3,2):[1]Sheet1!A3)', 5, [VALUE, { error: 'REF' }]],
  ['COUNTIF(G1:G3,[1]Sheet1!A1)', 2],
  ['SUMPRODUCT(--([1]Sheet1!B2:B4="x"),[1]Sheet1!C2:C4)', 40],
  ['COUNT([1]Sheet1!A1:C4)', 6],
  ['N([1]Sheet1!A2)', 2],
  ['ISNUMBER([1]Sheet1!A2)', true],
  ['LOOKUP(2,[1]Sheet1!A1:A3)', 2],
  ['HLOOKUP("v",[1]Sheet1!B1:C4,2,FALSE)', 10],
  ['CHOOSE(2,[1]Sheet1!A1,[1]Sheet1!A2)', 2],
  ['SUMIF(G1:G3,1,[1]Sheet1!A1:A3)', VALUE, [3]],
  ['COUNTIF([1]Sheet1!A1,1)', VALUE, [1]],
  ['_xlfn.XMATCH(3,[1]Sheet1!A1:A3)', 3],
  ['MATCH("x",[1]Sheet1!B:B,0)', 2],
  ['SUMIF(ExtB,"x",ExtC)', VALUE, [40]],
  ['SUM(ExtC)', 60],
]

const CASE: LinksCase = {
  id: 'functions',
  about: 'functions whose argument takes only a reference over a linked range, a linked criterion, a linked range through a name, more functions that read values',
  sheets: [{ name: 'S', cells: { F1: { value: 'k' }, F2: { value: 'x' }, G1: { value: 1 }, G2: { value: 1 }, G3: { value: 2 } } }],
  linked: { S: Object.fromEntries(FORMULAS.map(([f], i) => [`A${i + 1}`, f])) },
  names: [
    { name: 'ExtB', ref: '[1]Sheet1!$B$2:$B$4' },
    { name: 'ExtC', ref: '[1]Sheet1!$C$2:$C$4' },
  ],
  links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1, A2: 2, A3: 3, B1: 'k', C1: 'v', B2: 'x', C2: 10, B3: 'y', C3: 20, B4: 'x', C4: 30 } }] }],
  read: [{ sheet: 'S', range: `A1:A${FORMULAS.length}` }],
  predicted: Object.fromEntries(FORMULAS.map(([, v], i) => [`S!A${i + 1}`, v])),
  alternatives: Object.fromEntries(FORMULAS.flatMap(([, , alt], i) => (alt ? [[`S!A${i + 1}`, alt]] : []))),
}

export function buildLinksCases(): LinksCases {
  return { cases: [CASE] }
}
