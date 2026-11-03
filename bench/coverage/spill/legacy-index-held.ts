/**
 * Held out from `legacy-index` (`run.ts spill legacy-index-held`): names of literals, cells, other names and calls,
 * alone, as INDEX's place and inside calls, in formulas of the old kind. The text Excel shows is the converter's
 * (`SHOWN`), the values the engine's (`predicted`), both committed before Excel ran the set.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import PREDICTED from './legacy-index-held.predicted.json' with { type: 'json' }
import { fromLegacy } from '../../../src/xlsx.js'
import type { Val } from '../formulajs.js'

const DATA: Record<string, CellSpec> = Object.fromEntries([4, 8, 15, 16, 23, 42, -3.5, 10.25].map((v, i) => [`B${i + 1}`, v]))

const NAMES = Object.entries({
  q_neg: '-2',
  q_three: '3',
  q_true: 'TRUE',
  q_cell: "'S'!$B$3",
  q_two: '2',
  q_alias: 'q_two',
  q_err: '#N/A',
  q_box: "'S'!$B$2:$B$2",
  q_sum: "SUM('S'!$B$1:$B$8)",
  q_len: 'LEN("abc")',
  q_rng: "'S'!$B$1:$B$8",
}).map(([name, ref]) => ({ name, ref }))

const FORMULAS: Record<string, string> = {
  H1: '=q_neg',
  H2: '=INDEX(B1:B8,q_three)',
  H3: '=q_true',
  H4: '=INDEX(B1:B8,q_true)',
  H5: '=INDEX(B1:B8,q_cell)',
  H6: '=q_alias',
  H7: '=INDEX(B1:B8,q_alias)',
  H8: '=q_err',
  H9: '=q_box',
  H10: '=N(q_sum)',
  H11: '=IF(TRUE,q_sum)',
  H12: '=SUM(q_sum)',
  H13: '=q_len*2',
  H14: '=INDEX(q_rng,q_three)',
  H15: '=INDEX(q_rng,MATCH(8,q_rng,0))+q_len',
  H16: '=OFFSET(q_cell,1,0)',
  H17: '=OFFSET(q_rng,1,0,1,1)',
}

const at = Object.keys(FORMULAS)

export function buildSpillCases(): SpillCase[] {
  return [
    one('names', 'names of literals, cells, names and calls in formulas of the old kind', { ...DATA, ...FORMULAS }, at, {
      book: { sheets: [], names: NAMES },
      legacy: at.map((a) => `S!${a}`),
      predicted: (PREDICTED as Record<string, Record<string, Val>>).names,
    }),
  ]
}

/** The text Excel 365 was predicted to show for each formula (canonical, without `=`), committed before Excel ran. */
export const SHOWN: Readonly<Record<string, string>> = PREDICTED.shown

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = []

/** The converter's text of each formula, written into the prediction. */
export const converted = () => Object.fromEntries(at.map((a) => [`S!${a}`, fromLegacy(FORMULAS[a].slice(1), { names: NAMES, sheet: 'S' })]))
