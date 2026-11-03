/**
 * A linked book Excel can read, held out, on Windows alone (`bench/windows/links.ts`): the book is put beside the
 * opened one and opened, or left closed on the disk, then formulas over it are typed, read after a full calculation and
 * the book saved; against a link whose book is not there, which Excel marks (`refreshError`) as it cannot read it. The
 * book beside holds more than the link stores: a column, a text, and no cell Z99.
 *
 * `predicted` was committed before Excel ran the set. The guess: an open book is read as it is, a cell it does not hold
 * is blank and nothing is marked; a book closed after it was open leaves its values in the link, unmarked; a closed
 * book on the disk is read too when a formula over it is typed, but a formula typed before is not read again; a book
 * that is not there marks the link, and a cell it does not store is `#REF!`.
 */
import type { CachedValue, LinksCase, LinksCases, LinksOp } from './cases.js'

const REF = { error: 'REF' } as const

/** The book beside: Sheet1 with A1, a column H1:H3 and a text B5. */
const SOURCE = { Sheet1: { A1: 1, H1: 10, H2: 20, H3: 30, B5: 'x' } }

/** The formulas typed over `book`: a whole column, a cell the book does not hold, whether it is blank, a text. */
const typed = (book: string): LinksOp[] => [
  { op: 'enter', sheet: 'S', cell: 'A3', formula: `=SUM([${book}]Sheet1!H:H)` },
  { op: 'enter', sheet: 'S', cell: 'A4', formula: `=[${book}]Sheet1!Z99` },
  { op: 'enter', sheet: 'S', cell: 'A5', formula: `=ISBLANK([${book}]Sheet1!Z99)` },
  { op: 'enter', sheet: 'S', cell: 'A6', formula: `=[${book}]Sheet1!B5` },
]

/** A case over the link `book`, which stores A1 alone: the file's formulas read A1 and H2 (not stored), then `ops`. */
function over(id: string, about: string, book: string, ops: LinksOp[], predicted: Record<string, CachedValue | null>, alternatives: Record<string, (CachedValue | null)[]>): LinksCase {
  return {
    id,
    about,
    sheets: [{ name: 'S', cells: {} }],
    linked: { S: { A1: '[1]Sheet1!A1', A2: '[1]Sheet1!H2' } },
    links: [{ target: book, sheets: [{ name: 'Sheet1', cells: { A1: 1 } }] }],
    read: [{ sheet: 'S', range: 'A1:A6' }],
    ops,
    predicted,
    alternatives,
  }
}

const CASES: LinksCase[] = [
  over('open', 'the book opened beside, formulas typed over it, read and saved with it open', 'Open1.xlsx', [{ op: 'source', target: 'Open1.xlsx', open: true, sheets: SOURCE }, ...typed('Open1.xlsx')], { 'S!A1': 1, 'S!A2': 20, 'S!A3': 60, 'S!A4': 0, 'S!A5': true, 'S!A6': 'x' }, { 'S!A2': [0] }),
  over(
    'closed-after',
    'the book opened beside, formulas typed over it, then closed and removed before the full calculation',
    'Open2.xlsx',
    [{ op: 'source', target: 'Open2.xlsx', open: true, sheets: SOURCE }, ...typed('Open2.xlsx'), { op: 'dropSource', target: 'Open2.xlsx' }],
    { 'S!A1': 1, 'S!A2': 20, 'S!A3': 60, 'S!A4': 0, 'S!A5': true, 'S!A6': 'x' },
    { 'S!A2': [0, REF], 'S!A3': [REF], 'S!A4': [REF], 'S!A5': [false] },
  ),
  over(
    'disk',
    'the book beside, closed on the disk, formulas typed over it',
    'Open3.xlsx',
    [{ op: 'source', target: 'Open3.xlsx', open: false, sheets: SOURCE }, ...typed('Open3.xlsx')],
    { 'S!A1': 1, 'S!A2': 0, 'S!A3': 60, 'S!A4': 0, 'S!A5': true, 'S!A6': 'x' },
    { 'S!A2': [20, REF], 'S!A3': [REF], 'S!A4': [REF], 'S!A5': [false], 'S!A6': [REF] },
  ),
  over('missing', 'no book beside: formulas typed over the link', 'Open4.xlsx', typed('Open4.xlsx'), { 'S!A1': 1, 'S!A2': REF, 'S!A3': REF, 'S!A4': REF, 'S!A5': false, 'S!A6': REF }, { 'S!A2': [0] }),
]

export function buildLinksCases(): LinksCases {
  return { cases: CASES }
}
