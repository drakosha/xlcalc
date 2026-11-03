/**
 * Held out after `legacy-index-held-3` (`run.ts spill legacy-index-held-4`): INDEX over names, of references (a range,
 * a whole column, a union, a range by `:`) and of formulas (another name, OFFSET, INDEX, an expression, an array), with
 * one place and with two. The text Excel shows is the converter's (`SHOWN`), the values the engine's (`predicted`),
 * both committed before Excel ran.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import PREDICTED from './legacy-index-held-4.predicted.json' with { type: 'json' }
import { fromLegacy } from '../../../src/xlsx.js'
import type { Val } from '../formulajs.js'

const DATA: Record<string, CellSpec> = Object.fromEntries([4, 8, 15, 16, 23, 42, -3.5, 10.25].map((v, i) => [`B${i + 1}`, v]))

const NAMES = Object.entries({
  t_rng: "'S'!$B$1:$B$8",
  t_col: "'S'!$B:$B",
  t_nu: "('S'!$B$1:$B$3,'S'!$B$5:$B$8)",
  t_colon: "'S'!$B$1:'S'!$B$8",
  t_alias: 't_rng',
  t_off: "OFFSET('S'!$B$1,0,0,8)",
  t_idx: "INDEX('S'!$B$1:$C$8,0,1)",
  t_x: "'S'!$B$1:$B$8*1",
  t_arr: '{1,2;3,4}',
}).map(([name, ref]) => ({ name, ref }))

const FORMULAS: Record<string, string> = {
  H1: '=INDEX(t_rng,3)',
  H2: '=INDEX(t_col,3)',
  H3: '=INDEX(t_nu,2,1,2)',
  H4: '=INDEX(t_colon,3)',
  H5: '=INDEX(t_alias,3)',
  H6: '=INDEX(t_off,3)',
  H7: '=INDEX(t_idx,3)',
  H8: '=INDEX(t_x,MATCH(8,t_rng,0))',
  H9: '=INDEX(t_arr,2,1)',
  H10: '=INDEX(t_x,3,1)',
  H11: '=INDEX(t_rng,3,1)',
}

const at = Object.keys(FORMULAS)

export function buildSpillCases(): SpillCase[] {
  return [
    one('names', 'INDEX over names of references and of formulas, in formulas of the old kind', { ...DATA, ...FORMULAS }, at, {
      book: { sheets: [], names: NAMES },
      legacy: at.map((a) => `S!${a}`),
      predicted: PREDICTED.names as Record<string, Val>,
    }),
  ]
}

/** The text Excel 365 was predicted to show for each formula (canonical, without `=`), committed before Excel ran. */
export const SHOWN: Readonly<Record<string, string>> = PREDICTED.shown

/** The predictions Excel did not confirm: a name of `'S'!$B$1:'S'!$B$8` is a formula to INDEX, not a reference. */
export const MISSED: readonly string[] = ['S!H4']

/** The converter's text of each formula, written into the prediction. */
export const converted = () => Object.fromEntries(at.map((a) => [`S!${a}`, fromLegacy(FORMULAS[a].slice(1), { names: NAMES, sheet: 'S' })]))
