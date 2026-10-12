/**
 * Held out, predicted before Excel ran it by the grid's own model (`xlcalc-sheet/src/find-match.ts`): the query's
 * wildcards as the lookup functions read them, Look in Formulas the text a cell opens with for editing (`editText`),
 * Look in Values the text it shows. The questions: how Find reads `*`, `?` and `~` (a `~` before an ordinary character
 * or at the end, `?` over a character outside the BMP and over a combining mark, an empty query); which letters it
 * takes as one without Match case (Latin, Cyrillic, `ß`, the Turkish `İ` and `ı`); what Look in Formulas and Values read
 * in a formula, a formatted number, a long number, a boolean, an error, a text after an apostrophe, an empty text, the
 * cells of arrays, a merged area, a hidden row and column and a row a filter hid; the order by rows and by columns, in a
 * range and over the whole sheet. Both locales' texts are in the set, which runs in each.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, type FindBook, type FindCase, type FindQuery } from './cases.js'

const text = (...xs: string[]): Record<string, CellInput> => Object.fromEntries(xs.map((v, i) => [`A${i + 1}`, { value: v }]))

export const BOOKS: Record<string, FindBook> = {
  // the characters wildcards read; A11 is `e` and a combining acute accent, A12 the precomposed `é`
  wild: {
    cells: text('a*b', 'a?b', 'a~b', 'ab', 'axb', '~a', 'a', 'a~', 'xa', 'a😀b', 'éx', 'éx', '~', '*', '?', 'b'),
    columns: { A: { width: 20 } },
  },
  // letters in each case
  case: {
    cells: text('straße', 'STRASSE', 'STRAẞE', 'İstanbul', 'istanbul', 'ISTANBUL', 'ıstanbul', 'Привет', 'ПРИВЕТ', 'привет', 'ёлка', 'Ёлка', 'елка', 'Hello', 'HELLO', 'hello'),
    columns: { A: { width: 20 } },
  },
  // what Look in reads: values with formats, formulas, arrays, a merged area, a hidden row and column, rows a filter hid
  look: {
    cells: {
      D1: { value: 1 },
      D2: { value: 2 },
      B1: { formula: 'SUM(D1:D2)' },
      B2: { value: 0.05, numberFormat: '0%' },
      B3: { value: 1234.5, numberFormat: '"$"#,##0.00' },
      B4: { value: 46096, numberFormat: 14 },
      B5: { value: 1234567, numberFormat: '#,##0' },
      B6: { value: 123456789012345680 },
      B7: { value: true },
      B8: { formula: 'NA()' },
      B9: { formula: '1/0' },
      B10: { value: '007' },
      B11: { value: '007' },
      B12: { formula: '""' },
      F1: { formula: 'SEQUENCE(3)*7' },
      H1: { formula: 'D1:D2*5', array: 'H1:H2' },
      J1: { value: 'merged' },
      B19: { value: 'hiddenrow2' },
      B20: { value: 'hiddenrow' },
      L1: { value: 'hiddencol' },
      M1: { value: 'hiddencol2' },
      N30: { value: 'key' },
      N31: { value: 'keep' },
      N32: { value: 'drop' },
      N33: { value: 'keep' },
      B31: { value: 'filtered1' },
      B32: { value: 'filtered2' },
      B33: { value: 'filtered3' },
    },
    columns: { B: { width: 30 }, F: { width: 12 }, H: { width: 12 }, L: { hidden: true } },
    rows: { 20: { hidden: true } },
    merges: ['J1:K2'],
    filter: { range: 'N30:N33', field: 1, keep: 'keep' },
    // typed after an apostrophe: Excel keeps the apostrophe as the cell's prefix
    typed: { B11: "'007" },
  },
  // a 3×3 block with matches and three outside it
  order: {
    cells: {
      A1: { value: 'x1' }, B1: { value: 'y' }, C1: { value: 'x2' },
      A2: { value: 'x3' }, B2: { value: 'x4' }, C2: { value: 'z' },
      A3: { value: 'y' }, B3: { value: 'x5' }, C3: { value: 'x6' },
      E1: { value: 'x7' }, A5: { value: 'x8' }, E5: { value: 'x9' },
    },
  },
}

export function buildFindCases(): FindCase[] {
  const { out, add } = collector()
  const find = (group: string, book: string, what: string, more: Partial<FindQuery> = {}) => add(group, book, { what, lookIn: 'formulas', ...more })
  const whole = { whole: true }

  // `*` and `?`
  for (const what of ['*', 'a*b', '?']) find('wild', 'wild', what)
  for (const what of ['*', 'a*b', '*b', 'a*', '?', 'a?b', 'a??b', '?x', '??x', 'e?x']) find('wild', 'wild', what, whole)
  // `~` before a wildcard, before itself, before an ordinary character, at the end, alone
  for (const what of ['~*', '~?', '~~', '~a', 'a~', '~']) {
    find('tilde', 'wild', what)
    find('tilde', 'wild', what, whole)
  }
  for (const what of ['a~*b', 'a~?b', 'a~~b']) find('tilde', 'wild', what, whole)
  // an empty query and `*` over a range with empty cells, and an empty query over the whole sheet
  find('empty', 'wild', '', { range: 'A1:B4' })
  find('empty', 'wild', '', { range: 'A1:B4', ...whole })
  find('empty', 'wild', '*', { range: 'A1:B4' })
  find('empty', 'wild', '')
  // the same wildcards in values
  find('values', 'wild', 'a?b', { lookIn: 'values', ...whole })
  find('values', 'wild', '~*', { lookIn: 'values' })
  find('values', 'wild', '~a', { lookIn: 'values', ...whole })

  // the case of letters, with Match case and without
  for (const [what, matchCase] of [['hello', false], ['hello', true], ['HELLO', true], ['привет', false], ['привет', true], ['ПРИВЕТ', true], ['ё', false], ['Ё', false], ['ё', true], ['е', false], ['ß', false], ['ß', true], ['ẞ', false], ['ss', false], ['i', false], ['I', true], ['İ', false], ['ı', false]] as const)
    find('case', 'case', what, matchCase ? { matchCase } : {})
  for (const [what, matchCase] of [['hello', false], ['Hello', true], ['ПРИВЕТ', false], ['STRASSE', false], ['istanbul', false], ['İSTANBUL', false], ['h?llo', false], ['H*', true]] as const)
    find('case', 'case', what, { ...whole, ...(matchCase ? { matchCase } : {}) })

  // Look in: each query in formulas and in values
  const look = ['SUM', 'СУММ', '5%', '0.05', '0,05', '1,234', '1 234', '1 234', '1234.5', '1234,5', '$1', '2026', '46096', '3/15/2026', '15.03.2026', '1234567', 'E+17', '123456789012345', 'TRUE', 'ИСТИНА', '#N/A', '#Н/Д', 'NA()', 'НД()', '#DIV/0!', '#ДЕЛ/0!', '007', "'", '""', 'SEQUENCE', 'ПОСЛЕД', '14', '~*5', '10', 'merged', 'hiddenrow', 'hiddencol', 'filtered']
  for (const what of look) for (const lookIn of ['formulas', 'values'] as const) find('look', 'look', what, { lookIn })
  for (const lookIn of ['formulas', 'values'] as const) {
    // every cell Find takes as not empty, in a range and over the sheet
    find('look', 'look', '*', { lookIn, range: 'B1:B12' })
    find('look', 'look', '*', { lookIn })
  }
  for (const [what, lookIn] of [['5%', 'values'], ['0.05', 'formulas'], ['3/15/2026', 'formulas'], ['15.03.2026', 'formulas'], ['TRUE', 'formulas'], ['ИСТИНА', 'formulas']] as const) find('look', 'look', what, { lookIn, ...whole })

  // the order: by rows and by columns, over the sheet and in ranges, and a range of one cell
  for (const range of [undefined, 'A1:C3', 'B2:C3', 'B2'])
    for (const byColumns of [false, true]) find('order', 'order', 'x', { ...(range ? { range } : {}), ...(byColumns ? { byColumns } : {}) })
  return out
}
