/**
 * Held out after `legacy-index-held` (`run.ts spill legacy-index-held-2`): names of calls, other names and arrays in
 * the arguments that take a range's first cell, in IF and CHOOSE, and negative literals as INDEX's place. The text
 * Excel shows is the converter's (`SHOWN`), the values the engine's (`predicted`), both committed before Excel ran.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import PREDICTED from './legacy-index-held-2.predicted.json' with { type: 'json' }
import { fromLegacy } from '../../../src/xlsx.js'
import type { Val } from '../formulajs.js'

const DATA: Record<string, CellSpec> = Object.fromEntries([4, 8, 15, 16, 23, 42, -3.5, 10.25].map((v, i) => [`B${i + 1}`, v]))

const NAMES = Object.entries({
  r_len: 'LEN("abc")',
  r_arr: '{1,2}',
  r_txt: '"a"&"b"',
  r_rng: "'S'!$B$1:$B$8",
  r_alias: 'r_rng',
  r_neg: '-3',
  r_box: "'S'!$B$2:$B$2",
  r_two: '2',
  r_two2: 'r_two',
}).map(([name, ref]) => ({ name, ref }))

const FORMULAS: Record<string, string> = {
  H1: '=N(r_len)',
  H2: '=N(r_arr)',
  H3: '=T(r_txt)',
  H4: '=IF(TRUE,r_len)+0',
  H5: '=CHOOSE(1,r_len)',
  H6: '=r_alias',
  H7: '=SUM(r_alias)',
  H8: '=INDEX(B1:B8,r_neg)',
  H9: '=OFFSET(r_box,1,0)',
  H10: '=INDEX(B1:B8,r_two2)',
  H11: '=IFERROR(r_len,0)',
  H12: '=r_len&r_txt',
}

const at = Object.keys(FORMULAS)

export function buildSpillCases(): SpillCase[] {
  return [
    one('names', 'names of calls, names and arrays in more arguments, in formulas of the old kind', { ...DATA, ...FORMULAS }, at, {
      book: { sheets: [], names: NAMES },
      legacy: at.map((a) => `S!${a}`),
      predicted: PREDICTED.names as Record<string, Val>,
    }),
  ]
}

/** The text Excel 365 was predicted to show for each formula (canonical, without `=`), committed before Excel ran. */
export const SHOWN: Readonly<Record<string, string>> = PREDICTED.shown

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = []

/** The converter's text of each formula, written into the prediction. */
export const converted = () => Object.fromEntries(at.map((a) => [`S!${a}`, fromLegacy(FORMULAS[a].slice(1), { names: NAMES, sheet: 'S' })]))
