/**
 * What `values-held`'s two misses led to, held out (Excel en-US): FIXED and DOLLAR at the 255 characters TEXT keeps
 * (`FIXED(1E+200,0)` was #VALUE!), with commas, without them, a minus, a currency sign and parentheses; and the minus
 * of a conditional section before a text section (`[<0]0;@` showed -5 as `-5`, while `[<0]0;[>10]0;@` showed `5`): one
 * number section or two, other conditions and literals. Answers: `test/excel/cases/values-held-2.json`; `PREDICTED`,
 * the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, s, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['FIXED', 'LEN(FIXED(1E+191,0))', n(255)],
  ['FIXED', 'FIXED(1E+192,0)', e('VALUE')],
  ['FIXED', 'FIXED(-1E+191,0)', e('VALUE')],
  ['FIXED', 'FIXED(1E+191,1)', e('VALUE')],
  ['FIXED', 'LEN(FIXED(1E+254,0,TRUE))', n(255)],
  ['FIXED', 'FIXED(1E+255,0,TRUE)', e('VALUE')],
  ['FIXED', 'LEN(FIXED(-1E+253,0,TRUE))', n(255)],
  ['FIXED', 'FIXED(-1E+254,0,TRUE)', e('VALUE')],
  ['DOLLAR', 'LEN(DOLLAR(1E+190,0))', n(255)],
  ['DOLLAR', 'DOLLAR(1E+191,0)', e('VALUE')],
  ['DOLLAR', 'DOLLAR(-1E+190,0)', e('VALUE')],
  ['DOLLAR', 'LEN(DOLLAR(-1E+189,0))', e('VALUE')],
  ['TEXT', 'TEXT(-5,"[<0]0")', s('5')],
  ['TEXT', 'TEXT(-5,"[<=-1]0;@")', s('-5')],
  ['TEXT', 'TEXT(-5,"[<0]0.0;""x""@")', s('-5.0')],
  ['TEXT', 'TEXT(-5,"[<0]""neg ""0;@")', s('-neg 5')],
  ['TEXT', 'TEXT(-5,"[<0]0;0;@")', s('5')],
  ['TEXT', 'TEXT(-5,"[>-10]0;@")', s('-5')],
  ['TEXT', 'TEXT(5,"[<0]0;@")', e('VALUE')],
  ['TEXT', 'TEXT(-5,"[<0]0;[<-1]0;@")', s('5')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/** The prediction Excel did not confirm: DOLLAR(1E+190,0), 255 characters, is #VALUE!. */
export const MISSED: readonly string[] = ['DOLLAR#1']
