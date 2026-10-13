/**
 * Held out, predicted by the engine's `replaceInCells` before Excel ran it, after `rules-held-2` showed that a text put
 * back by Replace gets no percent editor (`-5` is -500%) and that a cell of the text format takes a number as a number.
 * The questions: whether a cell of the text format takes a formula, a boolean and a date too; a percent with a sign, with
 * a second `%`, a number by a percent made smaller; a date's separators, a money format's point; a number that becomes a
 * formula, a formula that becomes an error.
 */
import { collector } from '../find/cases.js'
import type { ReplaceBook, ReplaceCase, ReplaceQuery } from './cases.js'

export const BOOKS: Record<string, ReplaceBook> = {
  more: {
    cells: {
      A1: { value: '#1+1', numberFormat: '@' },
      A2: { value: 'xTRUE', numberFormat: '@' },
      A3: { value: 'x3/15/2026', numberFormat: '@' },
      A4: { value: 'x15.03.2026', numberFormat: '@' },
      B1: { value: 0.05, numberFormat: '0%' },
      B2: { value: 0.5, numberFormat: '0%' },
      B3: { value: 46096, numberFormat: 14 },
      B4: { value: 1234.5, numberFormat: '"$"#,##0.00' },
      B5: { value: 12 },
      B6: { formula: '1/1' },
    },
    columns: { A: { width: 20 }, B: { width: 20 } },
  },
}

export function buildReplaceCases(): ReplaceCase[] {
  const { out, add } = collector<ReplaceQuery>()
  const put = (what: string, by: string, range: string) => add('more', 'more', { what, by, range })
  put('#', '=', 'A1:A2')
  put('x', '', 'A2:A3')
  put('x', '', 'A4:A5')
  put('5', '+5', 'B1:B2')
  put('%', '%%', 'B1:B2')
  put('0', '', 'B2:B3')
  put('/', '-', 'B3:B4')
  put('.', '', 'B4:B5')
  put('1', '=1+', 'B5:B6')
  put('/1', '/0', 'B5:B6')
  return out
}
