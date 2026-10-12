/**
 * Held out, predicted by the engine's `replaceInCells` before Excel ran it, after `rules-held` showed that `*` takes as
 * little as it can (`a*b` in `a1b2b` is `a1b`), that a cell's text goes back as its own text edited (`6%` is 0.06 and `5`
 * 500% in a percent cell), that every cell found is entered again and an array of the old kind whole. The questions:
 * `*` at the start, in the middle, twice, before `~*` and at the end; whether a number begun by a sign gets the percent
 * editor's `%`; an apostrophe put, a text that becomes a formula that does not read, a formula that no longer reads
 * beside a text in a range of two areas, the list separator of the book's language, a formula emptied whole or of its
 * `=`, a string in a formula, a cell of the text format, `~` at the end of the query; whether a number too wide for its
 * column widens it; a table's header and the formulas that name its column.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector } from '../find/cases.js'
import type { ReplaceBook, ReplaceCase, ReplaceQuery } from './cases.js'

const text = (...xs: string[]): Record<string, CellInput> => Object.fromEntries(xs.map((v, i) => [`A${i + 1}`, { value: v }]))

export const BOOKS: Record<string, ReplaceBook> = {
  // where `*` stops
  star: { cells: text('a1b2b', 'b1b2b', 'a1b2c3c', 'a*b*c', 'xaya', 'ab ab'), columns: { A: { width: 30 } } },
  more: {
    cells: {
      A1: { value: 'x007' },
      A2: { value: '#1+' },
      A3: { value: 'x(y' },
      A4: { value: '12a', numberFormat: '@' },
      A5: { value: 'b~' },
      B1: { formula: 'SUM(D1,D2)' },
      B2: { value: -0.05, numberFormat: '0%' },
      B3: { formula: '1/0' },
      B4: { formula: '"a1"&A1' },
      B5: { value: 0.05, numberFormat: '0%' },
      C1: { value: 12345 },
      D1: { value: 1 },
      D2: { value: 2 },
      L1: { value: 'Qty' },
      M1: { value: 'Twice' },
      L2: { value: 2 },
      L3: { value: 3 },
      M2: { formula: '[@Qty]*2' },
      M3: { formula: '[@Qty]*2' },
    },
    columns: { A: { width: 20 }, B: { width: 20 } },
    tables: [{ name: 'T', range: 'L1:M3', columns: [{ name: 'Qty' }, { name: 'Twice', formula: '[@Qty]*2' }] }],
  },
}

export function buildReplaceCases(): ReplaceCase[] {
  const { out, add } = collector<ReplaceQuery>()
  for (const what of ['*b', 'a*', 'b*b', '?*b', 'a*b*c', '*~*', '**', 'a*a']) add('star', 'star', { what, by: 'x', range: 'A1:A6' })
  const more = (what: string, by: string, range?: string, whole?: true) => add('more', 'more', { what, by, ...(range ? { range } : {}), ...(whole ? { whole } : {}) })
  more('x', "'", 'A1:A2')
  more('#', '=', 'A1:A2')
  more('(', '', 'A3,B1')
  more('a', '', 'A4:A5')
  more('b~', 'z', 'A4:A5')
  more(',', ';', 'B1:B2')
  more(';', ',', 'B1:B2')
  more('%', '', 'B1:B2')
  more('5', '', 'B5:B6')
  more('=1/0', '', 'B3:B4', true)
  more('=', '', 'B3:B4')
  more('a1', 'b2', 'B4:B5')
  more('5', '56789012', 'C1:C2')
  more('Qty', 'Count')
  return out
}
