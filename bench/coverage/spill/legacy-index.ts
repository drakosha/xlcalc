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

/** Names of formulas: whether Excel puts `@` before the name, and whether a name stands for INDEX's row. */
const FORMULA_NAMES: Record<string, [string, string]> = {
  p_1: ['1+1', 'p_1'],
  p_2: ['ABS(-2)', 'p_2'],
  p_3: ['PI()', 'p_3'],
  p_4: ["SUM('S'!$B$1:$B$8)", '@p_4'],
  p_5: ["ROWS('S'!$B$1:$B$8)", '@p_5'],
  p_6: ["'S'!$B$2", 'p_6'],
  p_7: ["'S'!$B$2*2", 'p_7'],
  p_8: ["INDEX('S'!$B$1:$B$8,2)", '@p_8'],
  p_9: ["MATCH(8,'S'!$B$2,0)", 'p_9'],
  p_10: ['"x"', 'p_10'],
  p_11: ['MATCH(8,{4,8,15},0)', '@p_11'],
}

/** INDEX's row or column as a name, and the text predicted. */
const PLACE_NAMES: Record<string, [string, string]> = {
  K1: ['INDEX(B1:B8,z)', '@INDEX(B1:B8,z)'],
  K2: ['INDEX(B1:B8,half)', '@INDEX(B1:B8,half)'],
  K3: ['INDEX(B1:B8,par)', 'INDEX(B1:B8,par)'],
  K4: ['INDEX(B1:B8,p_1)', '@INDEX(B1:B8,p_1)'],
  K5: ['INDEX(B1:B8,p_9)', '@INDEX(B1:B8,p_9)'],
  K6: ['INDEX(B1:B8,txt)', '@INDEX(B1:B8,txt)'],
  K7: ['INDEX(blk,one,one)', 'INDEX(blk,one,one)'],
}

const NAMES = [
  { name: 'nm', ref: "'S'!$B$1:$B$8" },
  { name: 'blk', ref: "'S'!$B$1:$C$8" },
  { name: 'pos', ref: "MATCH(8,'S'!$B$1:$B$8,0)" },
  { name: 'two', ref: '2' },
]

export function buildSpillCases(): SpillCase[] {
  const ids = Object.keys(FORMULAS)
  const tops = Object.keys(FORMULA_NAMES).map((n, i) => [`J${i + 1}`, `=${n}`])
  const places = Object.entries(PLACE_NAMES).map(([a, [f]]) => [a, `=${f}`])
  const more = [...tops, ...places].map(([a]) => a)
  const names = [
    ...Object.entries(FORMULA_NAMES).map(([name, [ref]]) => ({ name, ref })),
    ...Object.entries({ z: '0', half: '2.5', par: '(2)', txt: '"2"', one: '1' }).map(([name, ref]) => ({ name, ref })),
    { name: 'blk', ref: "'S'!$B$1:$C$8" },
  ]
  return [
    one('names', 'INDEX over a name, with a name for its row', { ...DATA, ...FORMULAS }, ids, {
      book: { sheets: [], names: NAMES },
      legacy: ids.map((a) => `S!${a}`),
      predicted: (PREDICTED as Record<string, Record<string, Val>>).names,
    }),
    one('formulas', 'names of formulas, alone and as INDEX\'s row', { ...DATA, ...Object.fromEntries([...tops, ...places]) }, more, {
      book: { sheets: [], names },
      legacy: more.map((a) => `S!${a}`),
      predicted: (PREDICTED as Record<string, Record<string, Val>>).formulas,
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
  ...Object.fromEntries(Object.values(FORMULA_NAMES).map(([, t], i) => [`S!J${i + 1}`, t])),
  ...Object.fromEntries(Object.entries(PLACE_NAMES).map(([a, [, t]]) => [`S!${a}`, t])),
}

/**
 * The predictions Excel did not confirm: a name of any formula but a literal or a cell gets `@` (`1+1`, `(2)`, PI(),
 * MATCH over one cell too), and a name of a whole number literal from 1 is INDEX's place as the literal is.
 */
export const MISSED: readonly string[] = ['S!H2', 'S!H3', 'S!H6', 'S!J1', 'S!J2', 'S!J3', 'S!J7', 'S!J9', 'S!K3', 'S!K4', 'S!K5']
