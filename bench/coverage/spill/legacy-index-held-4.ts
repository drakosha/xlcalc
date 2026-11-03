/**
 * Held out after `legacy-index-held-3` (`run.ts spill legacy-index-held-4`): INDEX over names, of references (a range,
 * a whole column, a union, a range by `:`) and of formulas (another name, OFFSET, INDEX, an expression, an array), with
 * one place and with two. The text Excel shows is the converter's (`SHOWN`), the values the engine's (`predicted`),
 * both committed before Excel ran.
 */
import { legacyNames } from './legacy-names.js'
import PREDICTED from './legacy-index-held-4.predicted.json' with { type: 'json' }

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

export const { buildSpillCases, SHOWN, converted } = legacyNames('INDEX over names of references and of formulas, in formulas of the old kind', NAMES, FORMULAS, PREDICTED)

/** The predictions Excel did not confirm: a name of `'S'!$B$1:'S'!$B$8` is a formula to INDEX, not a reference. */
export const MISSED: readonly string[] = ['S!H4']
