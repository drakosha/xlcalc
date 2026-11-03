/**
 * INDEX with a computed row through the book's names in formulas of the old kind (`run.ts spill legacy-index`; the
 * legacy set `index-calls` has the rest): a name for the range, for the MATCH that computes the row, for a number.
 * The text Excel 365 shows for each was predicted before Excel ran the set (`SHOWN`), the values too (`predicted`).
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import PREDICTED from './legacy-index.predicted.json' with { type: 'json' }
import type { Val } from '../formulajs.js'

const DATA: Record<string, CellSpec> = Object.fromEntries([4, 8, 15, 16, 23, 42, -3.5, 10.25].map((v, i) => [`B${i + 1}`, v]))

const FORMULAS: Record<string, string> = {
  H1: '=INDEX(nm,MATCH(8,nm,0))',
  H2: '=INDEX(B1:B8,pos)',
  H3: '=INDEX(B1:B8,two)',
  H4: '=INDEX(nm,2)',
  H5: '=INDEX(B1:B8,MATCH(8,nm,0))',
  H6: '=INDEX(nm,pos)',
  H7: '=INDEX(blk,MATCH(8,nm,0),1)',
}

const NAMES = [
  { name: 'nm', ref: "'S'!$B$1:$B$8" },
  { name: 'blk', ref: "'S'!$B$1:$C$8" },
  { name: 'pos', ref: "MATCH(8,'S'!$B$1:$B$8,0)" },
  { name: 'two', ref: '2' },
]

export function buildSpillCases(): SpillCase[] {
  const ids = Object.keys(FORMULAS)
  return [
    one('names', 'INDEX over a name, with a name for its row', { ...DATA, ...FORMULAS }, ids, {
      book: { sheets: [], names: NAMES },
      legacy: ids.map((a) => `S!${a}`),
      predicted: (PREDICTED as Record<string, Record<string, Val>>).names,
    }),
  ]
}

/** The text Excel 365 was predicted to show for each formula (canonical, without `=`), committed before Excel ran. */
export const SHOWN: Readonly<Record<string, string>> = {
  'S!H1': 'INDEX(nm,MATCH(8,nm,0))',
  'S!H2': '@INDEX(B1:B8,pos)',
  'S!H3': '@INDEX(B1:B8,two)',
  'S!H4': 'INDEX(nm,2)',
  'S!H5': 'INDEX(B1:B8,MATCH(8,nm,0))',
  'S!H6': '@INDEX(nm,pos)',
  'S!H7': 'INDEX(blk,MATCH(8,nm,0),1)',
}

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = []
