/**
 * Held out, predicted by the engine's `findCells` with Find's rules before Excel ran it, after `rules-held` showed that
 * Find skips hidden rows and columns and gives the cells over the whole sheet by the blocks they leave, each block in
 * turn. The questions: the blocks by columns as well as by rows, a range with hidden lines in it, a range of several
 * areas in the order given and areas over each other; what Look in reads in a column too narrow for its value (`##`), a
 * text that reads as a formula, a time, a percent with a decimal, a negative number in parentheses, a number with
 * decimals of its format, a reference to the sheet, a text with a line break, a formula's text result; and the Greek
 * sigma, whose final form `ς` has the capital `Σ` as `σ` has. Both locales' texts are in the set, which runs in each.
 */
import type { CellInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { collector, type FindBook, type FindCase, type FindQuery } from './cases.js'

/** A 5×5 block, each cell `x` and its address. */
const block = (): Record<string, CellInput> => {
  const out: Record<string, CellInput> = {}
  for (let r = 0; r < 5; r++) for (let c = 0; c < 5; c++) out[formatAddress(r, c)] = { value: `x${formatAddress(r, c)}` }
  return out
}

export const BOOKS: Record<string, FindBook> = {
  // row 3 and column C hidden: four blocks of visible cells, and cells past the block on either side
  bands: { cells: { ...block(), G2: { value: 'xG2' }, B7: { value: 'xB7' } }, rows: { 3: { hidden: true } }, columns: { C: { hidden: true } } },
  // the order book of rules-held, for ranges of several areas
  areas: {
    cells: {
      A1: { value: 'x1' }, B1: { value: 'y' }, C1: { value: 'x2' },
      A2: { value: 'x3' }, B2: { value: 'x4' }, C2: { value: 'z' },
      A3: { value: 'y' }, B3: { value: 'x5' }, C3: { value: 'x6' },
      E1: { value: 'x7' }, A5: { value: 'x8' }, E5: { value: 'x9' },
    },
  },
  // what Look in reads; column D is too narrow for its values
  look: {
    cells: {
      D1: { value: 46096, numberFormat: 14 },
      D2: { value: 123456.78, numberFormat: '0.00' },
      B3: { value: '=1+2' },
      B4: { value: 0.7708333333333334, numberFormat: 'h:mm' },
      B5: { value: 0.125, numberFormat: '0.0%' },
      B6: { value: -5, numberFormat: '0;(0)' },
      B7: { value: 'ΟΣ' },
      B8: { value: 'ος' },
      B9: { value: 'οσ' },
      B10: { value: 1.5, numberFormat: '0.00' },
      B11: { formula: "'S'!B10*2" },
      B12: { value: 'a\nb' },
      B13: { formula: '"abc"&"def"' },
    },
    columns: { B: { width: 30 }, D: { width: 3 } },
  },
}

export function buildFindCases(): FindCase[] {
  const { out, add } = collector()
  const find = (group: string, book: string, what: string, more: Partial<FindQuery> = {}) => add(group, book, { what, lookIn: 'formulas', ...more })

  // the blocks hidden lines leave, by rows and by columns, over the sheet and in ranges
  for (const range of [undefined, 'A1:E5', 'B2:E5', 'A2:G7'])
    for (const byColumns of [false, true]) for (const lookIn of ['formulas', 'values'] as const) find('bands', 'bands', 'x', { lookIn, ...(range ? { range } : {}), ...(byColumns ? { byColumns } : {}) })

  // a range of several areas: in their order, the other order, and areas over each other
  for (const range of ['A1:A2,B2:C3', 'B2:C3,A1:A2', 'A1:B2,B2:C3', 'E5,A1:C1'])
    for (const byColumns of [false, true]) find('areas', 'areas', 'x', { range, ...(byColumns ? { byColumns } : {}) })

  // Look in: each query in formulas and in values
  for (const what of ['2026', '#', '123456', '=1', 'PM', '30', '12.5%', '12,5%', '0.125', '(5)', '-5', '1.5', '1,5', '1.50', '1,50', 'S!', 'cd'])
    for (const lookIn of ['formulas', 'values'] as const) find('look', 'look', what, { lookIn })
  find('look', 'look', 'a?b', { whole: true })
  find('look', 'look', 'a*b', { whole: true })
  for (const [what, matchCase] of [['σ', false], ['ς', false], ['Σ', true], ['ΟΣ', false]] as const) find('sigma', 'look', what, matchCase ? { matchCase } : {})
  return out
}
