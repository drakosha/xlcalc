/**
 * Held out, predicted before Excel ran it by the model of a grid's Replace All (`modelReplace`): Find's cells in
 * formulas, each one's text with the matches replaced, written as typed. The questions: what `*`, `?` and `~` match in
 * a replacement (`*` greedy, `?` over a character outside the BMP, half of one too), whether `*`, `?` and `~` in what is
 * put are taken as they are; a part and the whole cell, the case, several matches in a cell, an empty replacement; what
 * a cell turns into once its text is replaced: a text that reads as a number, a boolean, a date or a formula after it,
 * a text typed after an apostrophe, the same text put back; a formula (in the book's language, one that no longer reads,
 * one with an unknown function), a percent, a money format, a date, a boolean, an error, a spilling formula, an array of
 * the old kind whole and in part, a merged area; hidden rows and columns on a sheet with a filter and without one; a
 * range, a range of one cell. Both locales' texts are in the set, which runs in each.
 */
import type { CellInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { collector } from '../find/cases.js'
import type { ReplaceBook, ReplaceCase, ReplaceQuery } from './cases.js'

const text = (...xs: string[]): Record<string, CellInput> => Object.fromEntries(xs.map((v, i) => [`A${i + 1}`, { value: v }]))

export const BOOKS: Record<string, ReplaceBook> = {
  // the wildcards' texts; A10 a character outside the BMP alone
  wild: {
    cells: text('abc', 'a😀b', 'Hello hello HELLO', 'a1b2b', 'a*b', 'a?b', 'a~b', 'xyz', 'b', '😀'),
    columns: { A: { width: 30 } },
  },
  // texts that read as something else once replaced; A4 typed after an apostrophe
  entry: {
    cells: text('a1', '#1+1', '007', '007', 'x5%', 'TRUE x', 'ИСТИНА x', 'q 3/15/2026', 'q 15.03.2026', '1,5'),
    columns: { A: { width: 30 } },
    typed: { A4: "'007" },
  },
  // formulas, formats, arrays, a merged area, a hidden row and column
  look: {
    cells: {
      D1: { value: 1 },
      D2: { value: 2 },
      B1: { formula: 'SUM(D1:D2)' },
      B2: { value: 0.05, numberFormat: '0%' },
      B3: { value: 1234.5, numberFormat: '"$"#,##0.00' },
      B4: { value: 46096, numberFormat: 14 },
      B5: { value: 1234567, numberFormat: '#,##0' },
      B7: { value: true },
      B9: { formula: '1/0' },
      B12: { formula: 'D1+D1' },
      F1: { formula: 'SEQUENCE(3)*7' },
      H1: { formula: 'D1:D2*5', array: 'H1:H2' },
      J1: { value: 'merged' },
      B20: { value: 'hidden row' },
      L1: { value: 'hidden column' },
    },
    columns: { B: { width: 30 }, F: { width: 12 }, H: { width: 12 }, L: { hidden: true } },
    rows: { 20: { hidden: true } },
    merges: ['J1:K2'],
    read: ['F2', 'F3', 'H2', 'I1', 'K1'],
  },
}

/** A 5×8 block of `x` and its address, a row 10 hidden by hand past the filter's rows, column C hidden, the filter's key in column I hiding rows 3 and 6. */
BOOKS.filtered = (() => {
  const cells: Record<string, CellInput> = { I1: { value: 'key' }, A10: { value: 'xA10' } }
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 5; c++) cells[formatAddress(r, c)] = { value: `x${formatAddress(r, c)}` }
    if (r) cells[`I${r + 1}`] = { value: r + 1 === 3 || r + 1 === 6 ? 'drop' : 'keep' }
  }
  return { cells, rows: { 10: { hidden: true } }, columns: { C: { hidden: true } }, filter: { range: 'I1:I8', field: 1, keep: 'keep' } }
})()

export function buildReplaceCases(): ReplaceCase[] {
  const { out, add } = collector<ReplaceQuery>()
  const put = (group: string, book: string, what: string, by: string, more: Partial<ReplaceQuery> = {}) => add(group, book, { what, by, ...more })
  const wild = { range: 'A1:A10' }

  // `*`, `?` and `~` in the query
  for (const [what, by] of [['*', 'x'], ['a*b', 'x'], ['a?b', 'y'], ['a??b', 'y'], ['~*', '+'], ['~?', '!'], ['~~', '-']]) put('wild', 'wild', what, by, wild)
  // `?` over a character outside the BMP: each half, the first half
  put('wild', 'wild', '?', 'x', { range: 'A2:A3' })
  put('wild', 'wild', 'a?', 'x', { range: 'A2:A3' })
  put('wild', 'wild', '?', 'x', { ...wild, whole: true })
  put('wild', 'wild', '??', 'x', { ...wild, whole: true })
  // `*`, `?` and `~` in what is put
  for (const by of ['~*', '*?', '~']) put('put', 'wild', 'b', by, wild)
  // the whole cell, the case, several matches, nothing put
  put('whole', 'wild', 'abc', 'x', { ...wild, whole: true })
  put('whole', 'wild', 'a*', 'z', { ...wild, whole: true })
  put('whole', 'wild', 'b', 'x', { ...wild, whole: true })
  put('case', 'wild', 'hello', 'x', wild)
  put('case', 'wild', 'hello', 'x', { ...wild, matchCase: true })
  put('case', 'wild', 'HELLO', 'x', { ...wild, whole: true, matchCase: true })
  put('many', 'wild', 'b', 'XY', wild)
  put('many', 'wild', 'b', '', wild)
  put('many', 'wild', 'abc', '', { ...wild, whole: true })

  // what a text turns into: a number, a formula, a text that reads as a number, typed after an apostrophe, the same text
  const entry = (what: string, by: string, range: string) => put('entry', 'entry', what, by, { range })
  entry('a', '', 'A1:A2')
  entry('#', '=', 'A1:A2')
  entry('7', '8', 'A3:A4')
  entry('0', '0', 'A3:A4')
  entry('x', '', 'A5:A6')
  entry(' x', '', 'A6:A7')
  entry('q ', '', 'A8:A9')
  entry(',', ',', 'A10:A11')

  // formulas in the book's language, one that no longer reads, an unknown function, a reference in several places
  const look = (what: string, by: string, range?: string) => put('look', 'look', what, by, range ? { range } : {})
  look('SUM', 'MAX', 'B1:B2')
  look('СУММ', 'МАКС', 'B1:B2')
  look('SUM', 'SUMX', 'B1:B2')
  look('СУММ', 'СУММХ', 'B1:B2')
  look('(', '', 'B1:B2')
  look(':', '::', 'B1:B2')
  look('D1', 'D2', 'B12:B13')
  // a formula's value, which Replace does not read
  look('2', '9', 'B12:B13')
  // numbers by their formats, a date, a boolean, an error's formula
  look('5', '6', 'B2:B3')
  look('%', '', 'B2:B3')
  look('0', '1', 'B2:B3')
  look('4', '9', 'B3:B4')
  look('.', ',', 'B3:B4')
  look('2026', '2027', 'B4:B5')
  look('3', '4', 'B4:B5')
  look('TRUE', 'FALSE', 'B7:B8')
  look('ИСТИНА', 'ЛОЖЬ', 'B7:B8')
  look('RU', 'XX', 'B7:B8')
  look('0', '2', 'B9:B10')
  // a spilling formula, an array of the old kind whole and in part, a merged area
  look('7', '8', 'F1:F3')
  look('5', '6', 'H1:H2')
  look('5', '6', 'H1:I1')
  look('merged', 'm2', 'J1:K2')
  // the whole sheet: hidden rows and columns on a sheet without a filter, the formulas and the array together
  look('hidden', 'shown')
  look('D1', 'D2')
  // a range of one cell
  look('D1', 'D2', 'B12')

  // a sheet with a filter: hidden rows and columns, over the sheet and in a range
  put('filtered', 'filtered', 'x', 'y')
  put('filtered', 'filtered', 'x', 'y', { range: 'A1:E5' })
  put('filtered', 'filtered', 'x', 'y', { range: 'A9:E11' })
  return out
}
