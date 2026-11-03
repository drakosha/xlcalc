/**
 * Held out after `legacy-index-held` (`run.ts spill legacy-index-held-2`): names of calls, other names and arrays in
 * the arguments that take a range's first cell, in IF and CHOOSE, and negative literals as INDEX's place. The text
 * Excel shows is the converter's (`SHOWN`), the values the engine's (`predicted`), both committed before Excel ran.
 */
import { legacyNames } from './legacy-names.js'
import PREDICTED from './legacy-index-held-2.predicted.json' with { type: 'json' }

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

export const { buildSpillCases, SHOWN, converted } = legacyNames('names of calls, names and arrays in more arguments, in formulas of the old kind', NAMES, FORMULAS, PREDICTED)

/** The predictions Excel did not confirm: N takes a name of an array constant without `@` too (and gives #VALUE!): a name of any formula but a literal or a cell is a reference to Excel. */
export const MISSED: readonly string[] = ['S!H2']
