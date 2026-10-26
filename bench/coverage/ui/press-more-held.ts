/**
 * Held out, predicted by the rules `press-held` gave before Excel ran it: Increase and Decrease Decimal on General
 * numbers that General shows whole or not (the integer part long, many decimals, a sum of 0.1 and 0.2, in narrow and
 * wide columns), each also read without a press (what General shows); on codes of `#` alone, mixed `0.##`, sections
 * with digits in the zero section, a negative in parentheses, accounting 41 and 42 down. AutoSum past several blanks,
 * past a blank and text on the left, a blank above with numbers to the left, percents, an error and a boolean in the
 * run, one subtotal, a cell in row 1, numbers below the cell only.
 */
import type { CellInput, NumberFormat } from '../../../src/contract.js'
import { collector, type PressButton, type UiCase } from './cases.js'

const UP: PressButton = 'Increase Decimal'
const DOWN: PressButton = 'Decrease Decimal'

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const one = (group: string, value: CellInput['value'], numberFormat: NumberFormat | undefined, buttons: PressButton[], width?: number) =>
    add(group, { cells: { B2: { value, ...(numberFormat === undefined ? {} : { numberFormat }) } }, act: { do: 'press', range: 'B2', buttons }, ...(width === undefined ? {} : { columns: { B: { width } } }) })

  const numbers = [123456.789, 12345.6789, 1234.56789, 12345678.9, 98765432.1, 0.1 + 0.2, 2 / 3, 0.1234567891, 1234567.5, 3.14159265358979]
  for (const v of numbers) {
    one('shown', v, undefined, [])
    one('general', v, undefined, [UP])
    one('general', v, undefined, [DOWN])
  }
  for (const w of [5, 14]) for (const v of [12345.6789, 2 / 3]) {
    one('shown-width', v, undefined, [], w)
    one('general-width', v, undefined, [UP], w)
  }
  const codes: NumberFormat[] = ['#', '0.##', '#,##0.0#', '0.00;-0.00;0.00', '0.00_);(0.00)', '0.00 "kg";-0.00 "kg"', '#,##0;[Red]-#,##0', 41, 42]
  for (const f of codes) for (const b of [UP, DOWN]) one('code', 1.5, f, [b])

  const n = (value: number, numberFormat?: NumberFormat): CellInput => ({ value, ...(numberFormat === undefined ? {} : { numberFormat }) })
  const sum = (group: string, cells: Record<string, CellInput>, at: string) => add(group, { cells, act: { do: 'press', range: at, buttons: ['AutoSum'], enter: true } })
  sum('sum-far', { A1: n(1), A2: n(2) }, 'A9')
  sum('sum-left-gap', { A2: n(1), B2: n(2) }, 'D2')
  sum('sum-left-text', { A2: n(1), B2: n(2), C2: { value: 'x' } }, 'D2')
  sum('sum-above-blank', { D1: { value: 'h' }, A2: n(1), B2: n(2), C2: n(3) }, 'D2')
  sum('sum-percent', { A1: n(0.1, '0%'), A2: n(0.2, '0%') }, 'A3')
  sum('sum-error', { A1: n(1), A2: { formula: '1/0' }, A3: n(3) }, 'A4')
  sum('sum-boolean', { A1: n(1), A2: { value: true }, A3: n(3) }, 'A4')
  sum('sum-subtotal', { A1: n(1), A2: n(2), A3: { formula: 'SUM(A1:A2)' }, A4: n(4) }, 'A5')
  sum('sum-row-1', { A1: n(1), B1: n(2) }, 'C1')
  sum('sum-below', { A3: n(1), A4: n(2) }, 'A2')
  return out
}
