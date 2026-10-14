/**
 * Held out again, predicted by `refAt` before Excel ran it (`POINT_BY`): its rule was revised after `point-more-held`
 * (Excel enters a formula after an operand only when it reads as it stands, in the locale it is typed in), so 20 places
 * no set had: whole columns and rows and their halves, a spill reference, a sheet with a column alone, operators doubled,
 * a percent twice, nested brackets, numbers apart by a space and by `;`. Each a text typed over C3, Right and Return;
 * both locales. Typed on the keyboard: ASCII only.
 */
import { typedThenRight, type UiCase } from './cases.js'

export const POINT_BY = 'refAt'

const TEXTS = [
  '=A:A', '=1:1', '=A1:A', '=SUM(A1:', '=E2#', '=Sheet1!A', '=1.5.', '=-(', '=1%%', '=(1)', '=((', '=1+-', '=A1^2', '=TRUE+',
  '=@A1', '=-+', '=1 2', '=1;2', '="a"&"b"', '=A1:B',
]

export const buildUiCases = (): UiCase[] => typedThenRight(TEXTS)
