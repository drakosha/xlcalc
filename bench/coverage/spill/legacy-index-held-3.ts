/**
 * Held out after `legacy-index-held-2` (`run.ts spill legacy-index-held-3`): names of arrays, of ROW over a range and of
 * expressions over a range, which Excel takes as references, in the arguments that pass an array on, take one or take
 * an array, and as INDEX's array. The text Excel shows is the converter's (`SHOWN`), the values the engine's
 * (`predicted`), both committed before Excel ran.
 */
import { legacyNames } from './legacy-names.js'
import PREDICTED from './legacy-index-held-3.predicted.json' with { type: 'json' }

const NAMES = Object.entries({
  s_arr: '{4,5,6}',
  s_col: '{4;5;6}',
  s_rows: "ROW('S'!$B$1:$B$3)",
  s_x3: "'S'!$B$1:$B$8*3",
  s_txt: '{"a","b"}',
}).map(([name, ref]) => ({ name, ref }))

const FORMULAS: Record<string, string> = {
  H1: '=IFERROR(s_arr,0)',
  H2: '=INDEX(s_arr,2)',
  H3: '=INDEX(s_col,2)',
  H4: '=ROWS(s_col)',
  H5: '=s_rows',
  H6: '=SUM(s_rows)',
  H7: '=s_rows+0',
  H8: '=IFERROR(s_x3,0)',
  H9: '=T(s_txt)',
  H10: '=INDEX(s_x3,2)',
  H11: '=ABS(s_arr)',
  H12: '=SUMPRODUCT(s_x3)',
}

export const { buildSpillCases, SHOWN, converted } = legacyNames('names of arrays and expressions over ranges as references, in formulas of the old kind', NAMES, FORMULAS, PREDICTED)

/** The predictions Excel did not confirm: INDEX takes a name of a formula that is not a reference as an array, so with one place it gets `@`. */
export const MISSED: readonly string[] = ['S!H2', 'S!H3', 'S!H10']
