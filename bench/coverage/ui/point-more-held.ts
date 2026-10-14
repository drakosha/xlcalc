/**
 * Held out, predicted by `refAt` before Excel ran it (`POINT_BY`): after `point-held` gave the rule, more places of an
 * arrow pressed in text typed over C3 (Enter mode), each a text typed, Right and Return. A number's point and the
 * decimal separator after a number, a reference or a name; a closed string, a percent, an error literal; inside an
 * array constant; `$` alone and after a column; a sheet's `!`, quoted and not; an open table bracket; a closing bracket
 * too many; an exponent's `e` and its sign; spaces after an operator and after an operand; operators after operators.
 * Both locales' separators are in the set, which runs in each. Typed on the keyboard: ASCII only.
 */
import { typedThenRight, type UiCase } from './cases.js'

export const POINT_BY = 'refAt'

const TEXTS = [
  '=1.5', '=A1.', '=x.', '=1,', '=1,5', '="a"', '=5%', '=#N/A', '={1', '={1,', '={1;', '=$', '=A$', '=Sheet1!', "='My S'!", '=T[',
  '=A1)', '=1e', '=1e+', '=A1 + ', '=SUM(A1 ', '=-1', '= ', '=SUM(,', '==', '=A1<>', '=@',
]

export const buildUiCases = (): UiCase[] => typedThenRight(TEXTS)
