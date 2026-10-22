/**
 * The values r2/values changed without a recorded answer, held out (Excel en-US): FIXED and DOLLAR from 1E+21 on, digit
 * formats past the 15th significant digit below 1E+21 (TEXT, and FIXED of the same rule), a text section ending a code
 * of two or three sections (`h:mm;@`, `0;@`) with negative numbers and zero, and conditional codes ending in `;@`.
 * Answers: `test/excel/cases/values-held.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, s, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['FIXED', 'FIXED(1E+21)', s('1,000,000,000,000,000,000,000.00')],
  ['FIXED', 'FIXED(1E+21,2,TRUE)', s('1000000000000000000000.00')],
  ['FIXED', 'FIXED(-1E+21,0)', s('-1,000,000,000,000,000,000,000')],
  ['FIXED', 'FIXED(1E+21,-3)', s('1,000,000,000,000,000,000,000')],
  ['FIXED', 'FIXED(123456789012345678901234,0)', s('123,456,789,012,345,000,000,000')],
  ['FIXED', 'FIXED(2^70,1)', s('1,180,591,620,717,410,000,000.0')],
  ['FIXED', 'FIXED(2^60,0)', s('1,152,921,504,606,850,000')],
  ['FIXED', 'FIXED(1/3,20)', s('0.33333333333333300000')],
  ['FIXED', 'LEN(FIXED(1E+200,0))', n(267)],
  ['FIXED', 'LEN(FIXED(1E+200,0,TRUE))', n(201)],
  ['DOLLAR', 'DOLLAR(1E+21)', s('$1,000,000,000,000,000,000,000.00')],
  ['DOLLAR', 'DOLLAR(-1E+21)', s('($1,000,000,000,000,000,000,000.00)')],
  ['DOLLAR', 'DOLLAR(2^70,0)', s('$1,180,591,620,717,410,000,000')],
  ['TEXT', 'TEXT(1/3,"0.00000000000000000000")', s('0.33333333333333300000')],
  ['TEXT', 'TEXT(2^60,"0")', s('1152921504606850000')],
  ['TEXT', 'TEXT(2^60,"#,##0")', s('1,152,921,504,606,850,000')],
  ['TEXT', 'TEXT(793853.646,"0.0000000000")', s('793853.6460000000')],
  ['TEXT', 'TEXT(0.1+0.2,"0.00000000000000000")', s('0.30000000000000000')],
  ['TEXT', 'TEXT(1/3*1E+6,"#,##0.0000000000")', s('333,333.3333333330')],
  ['TEXT', 'TEXT(123456789012345000,"0%")', s('12345678901234500000%')],
  ['TEXT', 'TEXT(2^60,"# ?/?")', s('1152921504606850000    ')],
  ['TEXT', 'TEXT(-0.25,"h:mm;@")', e('VALUE')],
  ['TEXT', 'TEXT(-1,"m/d/yyyy;@")', e('VALUE')],
  ['TEXT', 'TEXT(-1,"m/d/yy;@")', e('VALUE')],
  ['TEXT', 'TEXT(-0.25,"[$-409]h:mm AM/PM;@")', e('VALUE')],
  ['TEXT', 'TEXT(1.5,"h:mm;@")', s('12:00')],
  ['TEXT', 'TEXT(-5,"0;@")', s('-5')],
  ['TEXT', 'TEXT(-5,"0.00;""t""@")', s('-5.00')],
  ['TEXT', 'TEXT(0,"0.00;-0.00;@")', s('0.00')],
  ['TEXT', 'TEXT(-5,"0;0;@")', s('5')],
  ['TEXT', 'TEXT(-5,"[>10]0;@")', e('VALUE')],
  ['TEXT', 'TEXT(5,"[<0]0;[>10]0;@")', e('VALUE')],
  ['TEXT', 'TEXT(-5,"[<0]0;[>10]0;@")', s('5')],
  ['TEXT', 'TEXT(20,"[<0]0;[>10]0;@")', s('20')],
  ['TEXT', 'TEXT(-5,"[<0]0;@")', s('5')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/** The predictions Excel did not confirm: FIXED past 255 characters is #VALUE!, and `[<0]0;@` shows -5 as `-5`. */
export const MISSED: readonly string[] = ['FIXED#9', 'TEXT#22']
