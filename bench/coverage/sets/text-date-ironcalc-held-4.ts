/**
 * TEXT's date and time codes by Microsoft's documentation, after `text-date-ironcalc-held-3`'s misses: the codes the
 * support pages "Number format codes" and "Review guidelines for customizing a number format" show (fractions of a
 * second in `h:mm:ss.00` and `[ss].00`, elapsed time, AM/PM and A/P, minutes after hours and before seconds), the
 * grammar of [MS-XLS] 2.4.126 (ECMA-376 18.8.31's codes: a point and one to three zeros anywhere after a section's
 * first date code, no other digit placeholder, percent or exponent there), and the edges they do not settle: minutes
 * before decimals across a space, decimals of two lengths, a quoted or escaped zero. Answers:
 * `test/excel/cases/text-date-ironcalc-held-4.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, s } from '../vals.js'

/** 2012-09-29 12:01:30.5 and 2012-03-05 07:45:12.25 */
const T = '(41181.5+90.5/86400)'
const U = '(40973+27912.25/86400)'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // the documentation's codes
  ['TEXT', `TEXT(${T},"h:mm:ss.00")`, s('12:01:30.50')],
  ['TEXT', `TEXT(${T},"[ss].00")`, s('3558081690.50')],
  ['TEXT', `TEXT(${T},"[h]:mm")`, s('988356:01')],
  ['TEXT', `TEXT(${T},"[mm]:ss")`, s('59301361:30')],
  ['TEXT', `TEXT(${T},"h:mm:ss A/P")`, s('12:01:30 P')],
  ['TEXT', `TEXT(${U},"h:mm AM/PM")`, s('7:45 AM')],
  ['TEXT', `TEXT(${U},"mm")`, s('03')],
  ['TEXT', `TEXT(${U},"hh:mm")`, s('07:45')],
  ['TEXT', `TEXT(${U},"mm:ss")`, s('45:12')],
  // the grammar: decimals anywhere after the first date code, one to three zeros, nothing else of a number
  ['TEXT', `TEXT(${U},"yyyy-mm-dd hh:mm:ss.000")`, s('2012-03-05 07:45:12.250')],
  ['TEXT', `TEXT(${U},"dd .000")`, s('05 .250')],
  ['TEXT', `TEXT(${U},"[h].000")`, s('983359.250')],
  ['TEXT', `TEXT(${U},"s.00 h")`, s('12.25 7')],
  ['TEXT', `TEXT(${U},"AM/PM h")`, s('AM 7')],
  ['TEXT', `TEXT(${U},"ss.0000")`, e('VALUE')],
  ['TEXT', `TEXT(${U},".00 dd")`, e('VALUE')],
  ['TEXT', `TEXT(${U},"dd%")`, e('VALUE')],
  ['TEXT', `TEXT(${U},"dd E+00")`, e('VALUE')],
  ['TEXT', `TEXT(${U},"dd ?")`, e('VALUE')],
  // edges the documentation leaves open
  ['TEXT', `TEXT(${U},"m.00")`, s('45.25')],
  ['TEXT', `TEXT(${U},"mm .00")`, s('45 .25')],
  ['TEXT', `TEXT(${U},"mm/dd .00")`, s('03/05 .25')],
  ['TEXT', `TEXT(${U},"ss.0 .000")`, s('12.2 .250')],
  ['TEXT', `TEXT(${U},"h.0 AM/PM")`, s('7.2 AM')],
  ['TEXT', `TEXT(${U},"dd ""0""")`, s('05 0')],
  ['TEXT', `TEXT(${U},"dd\\0")`, s('050')],
  // IronCalc's code again
  ['TEXT', `TEXT(${U},"mm###")`, e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
