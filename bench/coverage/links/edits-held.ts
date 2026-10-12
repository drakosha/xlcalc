/**
 * Links to other workbooks after edits, held out. In `basics-held` a cell the link holds no value of was blank after
 * opening, and `#REF!` after a fill down, a copy and a sheet renamed with a full calculation. Each case here takes one
 * edit of this book on the same link, then a full calculation, to tell which edit does it: none; a fill down whose new
 * references the link holds; one whose new reference it does not; a copy to a place whose reference the link holds and
 * one it does not; this book's sheet named like the link's sheet renamed; rows inserted on another sheet.
 *
 * `predicted` was committed before Excel ran the set. The guess: a new reference to a cell the link holds no value of
 * makes every such cell `#REF!` at the next calculation (the link's values are then taken as not complete); an edit
 * that adds none leaves them blank. `alternatives` holds the other answer.
 */
import type { CachedValue, LinkedBook, LinksCase, LinksCases, LinksOp } from './cases.js'

const BOOK2: LinkedBook = { target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1, A2: 2, A3: 3, B1: 'b' } }] }

/** S!A1 a cell the link holds no value of, A2 a whole column, A3 a cell it holds; the edit's formulas in column C. */
const FORMULAS = { A1: '[1]Sheet1!A9', A2: 'SUM([1]Sheet1!A:A)', A3: '[1]Sheet1!A1', C1: '[1]Sheet1!A1', C5: '[1]Sheet1!A3' }

function edit(id: string, about: string, ops: LinksOp[], blank: boolean, more: Record<string, CachedValue | null> = {}): LinksCase {
  const lost: CachedValue = { error: 'REF' }
  return {
    id,
    about,
    sheets: [
      { name: 'S', cells: {} },
      { name: 'Sheet1', cells: { A1: { value: 1000 } } },
      { name: 'T', cells: { A1: { value: 1 } } },
    ],
    linked: { S: FORMULAS, T: { A3: '[1]Sheet1!A2' } },
    links: [BOOK2],
    read: [
      { sheet: 'S', range: 'A1:A3' },
      { sheet: 'S', range: 'C1:D10' },
    ],
    ops,
    predicted: { 'S!A1': blank ? 0 : lost, 'S!A2': blank ? 6 : lost, 'S!A3': 1, ...more },
    alternatives: { 'S!A1': [blank ? lost : 0], 'S!A2': [blank ? lost : 6] },
  }
}

const CASES: LinksCase[] = [
  edit('none', 'no edit, a full calculation', [], true),
  edit('fill-held', 'a fill down whose new references the link holds', [{ op: 'fillDown', sheet: 'S', to: 'C1:C3' }], true, { 'S!C2': 2, 'S!C3': 3 }),
  edit('fill-lost', 'a fill down whose new reference the link does not hold', [{ op: 'fillDown', sheet: 'S', to: 'C5:C6' }], false, { 'S!C6': { error: 'REF' } }),
  edit('copy-held', 'a copy whose new reference the link holds', [{ op: 'copyRange', sheet: 'S', range: 'C1', to: 'D1' }], true, { 'S!D1': 'b' }),
  edit('copy-lost', 'a copy whose new reference the link does not hold', [{ op: 'copyRange', sheet: 'S', range: 'C1', to: 'D10' }], false, { 'S!D10': { error: 'REF' } }),
  edit('rename', 'this book’s sheet named like the link’s sheet renamed', [{ op: 'renameSheet', sheet: 'Sheet1', to: 'Local1' }], true),
  edit('insert', 'rows inserted on another sheet above a formula with a link', [{ op: 'insertRows', sheet: 'T', range: 'A1' }], true),
]

export function buildLinksCases(): LinksCases {
  return { cases: CASES }
}
