/**
 * TEXT's point and zeros after a date code, as `text-date-ironcalc-held-2`'s misses showed it: the second's decimals
 * after a day, a year, a month's name, hours, elapsed hours and seconds, with `m` before them minutes; and what is no
 * such point: a space between, `#` after a zero, a second point, after AM/PM. Answers:
 * `test/excel/cases/text-date-ironcalc-held-3.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, s } from '../vals.js'

/** 2012-09-29 12:01:30.5 */
const T = '(41181.5+90.5/86400)'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['TEXT', `TEXT(${T},"d.00")`, s('29.50')],
  ['TEXT', `TEXT(${T},"yyyy.0")`, s('2012.5')],
  ['TEXT', `TEXT(${T},"mmmm.00")`, s('September.50')],
  ['TEXT', `TEXT(${T},"ddd.000")`, s('Sat.500')],
  ['TEXT', `TEXT(${T},"h.00")`, s('12.50')],
  ['TEXT', `TEXT(${T},"[h].00")`, s('988356.50')],
  ['TEXT', `TEXT(${T},"[ss].0")`, s('3558081690.5')],
  ['TEXT', `TEXT(${T},"mm.00 ss")`, s('01.50 30')],
  ['TEXT', `TEXT(${T},"dd .00")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"dd.0#")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"yy.00.00")`, e('VALUE')],
  ['TEXT', `TEXT(${T},"h AM/PM.00")`, e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
