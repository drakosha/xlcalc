/**
 * Held out again, predicted by `refAt` before Excel ran it (`POINT_BY`): its rule was revised after `point-more-held`
 * (Excel enters a formula after an operand only when it reads as it stands, in the locale it is typed in), so 20 places
 * no set had: whole columns and rows and their halves, a spill reference, a sheet with a column alone, operators doubled,
 * a percent twice, nested brackets, numbers apart by a space and by `;`. Each a text typed over C3, Right and Return;
 * both locales. Typed on the keyboard: ASCII only.
 */
import { collector, type UiCase } from './cases.js'
import type { CellInput } from '../../../src/contract.js'

export const POINT_BY = 'refAt'

const CELLS: Record<string, CellInput> = { A1: { value: 2 }, D3: { value: 10 }, D4: { value: 20 } }

const TEXTS = [
  '=A:A', '=1:1', '=A1:A', '=SUM(A1:', '=E2#', '=Sheet1!A', '=1.5.', '=-(', '=1%%', '=(1)', '=((', '=1+-', '=A1^2', '=TRUE+',
  '=@A1', '=-+', '=1 2', '=1;2', '="a"&"b"', '=A1:B',
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  for (const text of TEXTS) add('typed', { cells: CELLS, act: { do: 'point', at: 'C3', steps: [{ type: text }, { key: 'right' }, { key: 'return' }] }, read: ['C4'] })
  return out
}
