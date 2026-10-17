/**
 * TEXT's `m` before a point and zeros, as `text-date-ironcalc-held`'s miss showed it: minutes, with the decimals of
 * the second (90.5 seconds past noon tell them from a fraction of the minute), in `m.00`, `mm.000`, after hours; and
 * the codes it is not: `mmm`, years, days, elapsed minutes, `#` after the point. Answers:
 * `test/excel/cases/text-date-ironcalc-held-2.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, s } from '../vals.js'

/** 2012-09-29 12:01:30.5 */
const T = '(41181.5+90.5/86400)'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['TEXT', `TEXT(${T},"mm.00")`, s('01.50')],
  ['TEXT', `TEXT(${T},"m.00")`, s('1.50')],
  ['TEXT', `TEXT(${T},"mm.000")`, s('01.500')],
  ['TEXT', `TEXT(${T},"mm.0")`, s('01.5')],
  ['TEXT', `TEXT(${T},"hh:mm.00")`, s('12:01.50')],
  ['TEXT', `TEXT(${T},"yyyy-mm.00")`, s('2012-01.50')],
  ['TEXT', `TEXT(${T},"mm.##")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"mmm.00")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"yy.00")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"dd.00")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"[mm].00")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"mm.00 dd")`, s('01.50 29')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
