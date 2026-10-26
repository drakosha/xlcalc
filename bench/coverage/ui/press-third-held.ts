/**
 * Held out, predicted by the rules `press-held` and `press-more-held` gave before Excel ran it: Increase and Decrease
 * Decimal on General fractions the column shows as whole numbers (negative, more decimals, in a wider column that
 * shows them) and on numbers the column shows as `#` (narrow columns, small and large numbers); AutoSum over two
 * blocks of subtotals, a subtotal nearest with a block below it, a subtotal written with another function, and a
 * subtotal on the left.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, type PressButton, type UiCase } from './cases.js'

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const one = (group: string, value: number, buttons: PressButton[], width?: number) =>
    add(group, { cells: { B2: { value } }, act: { do: 'press', range: 'B2', buttons }, ...(width === undefined ? {} : { columns: { B: { width } } }) })
  for (const v of [-12345678.9, 99999999.25, 1234567.125]) for (const b of ['Increase Decimal', 'Decrease Decimal'] as const) one('whole', v, [b])
  one('whole-wide', 12345678.9, ['Increase Decimal'], 12)
  for (const [v, w] of [[123456789, 4], [0.000123, 2], [98765.4, 3]] as const) for (const b of ['Increase Decimal', 'Decrease Decimal'] as const) one('hashes', v, [b], w)

  const n = (value: number): CellInput => ({ value })
  const f = (formula: string): CellInput => ({ formula })
  const sum = (group: string, cells: Record<string, CellInput>, at: string) => add(group, { cells, act: { do: 'press', range: at, buttons: ['AutoSum'], enter: true } })
  sum('sum-blocks', { A1: n(1), A2: f('SUM(A1)'), A3: n(3), A4: f('SUM(A3)'), A5: n(5), A6: n(6) }, 'A7')
  sum('sum-blocks', { A1: n(1), A2: f('SUM(A1)'), A3: n(3), A4: f('SUM(A3)') }, 'A5')
  sum('sum-other', { A1: n(1), A2: n(2), A3: f('SUBTOTAL(9,A1:A2)'), A4: n(4) }, 'A5')
  sum('sum-left', { A2: n(1), B2: f('SUM(A2)'), C2: n(3) }, 'D2')
  return out
}
