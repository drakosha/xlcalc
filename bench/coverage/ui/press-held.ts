/**
 * Held out, predicted by `stepDecimals` and `Session.getAutoSum` before Excel ran it: the Home tab's buttons pressed
 * with a range selected. Increase and Decrease Decimal on General cells of numbers that show decimals or none, in
 * columns of other widths, on text, a blank, a boolean and a date, pressed again; on codes of each kind (decimals, a
 * group separator, percent, scientific, accounting, currency, sections, colours, literals, the most decimals, dates,
 * times, fractions, text); on a selection of cells of other formats (the active cell's result for all). Comma Style
 * and Accounting Number Format on a number, a formatted number and text. AutoSum on one cell: numbers above under a
 * header and without one, with a blank in them or right above, to the left, above and to the left, none, text above,
 * subtotals of SUM above, formulas and dates above. The set runs in en-US: it presses the buttons by their names there.
 */
import type { CellInput, NumberFormat } from '../../../src/contract.js'
import { collector, type PressButton, type UiCase } from './cases.js'

const UP: PressButton = 'Increase Decimal'
const DOWN: PressButton = 'Decrease Decimal'

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const press = (group: string, cells: Record<string, CellInput>, range: string, buttons: PressButton[], extra: Partial<UiCase> = {}) =>
    add(group, { cells, act: { do: 'press', range, buttons }, ...extra })
  const one = (group: string, value: CellInput['value'] | undefined, numberFormat: NumberFormat | undefined, buttons: PressButton[], width?: number) =>
    press(group, { B2: { ...(value === undefined ? {} : { value }), ...(numberFormat === undefined ? {} : { numberFormat }) } }, 'B2', buttons, width === undefined ? {} : { columns: { B: { width } } })

  // General: the decimals the number shows, one more or one less
  for (const v of [1.5, 5, 1 / 3, 1234567.891, 1e20, 0.000012, -2.25]) for (const b of [UP, DOWN]) one('general', v, undefined, [b])
  // in a wide and a narrow column
  for (const w of [20, 4]) one('general-width', 1 / 3, undefined, [UP], w)
  // no number: text, a blank, a boolean; a date in its format; pressed three times
  for (const v of ['abc', undefined, true]) for (const b of [UP, DOWN]) one('general-other', v, undefined, [b])
  one('date', 45688, 14, [UP])
  one('again', 1.5, undefined, [UP, UP, UP])
  one('again', 1.5, '0.00', [DOWN, DOWN, DOWN])
  // codes of each kind, a step each way
  const codes: NumberFormat[] = [
    '0.00', '0.0', '0', '#,##0', '#,##0.00', 44, 42, 43, 41, '0.00E+00', '0E+00', '0%', '0.0%', '$#,##0.00', '"$"#,##0.00_);[Red]("$"#,##0.00)',
    '#,##0.00 [$₽-419]', '0.00;-0.00;"zero"', '[Red]0.00', '0.00 "kg"', 'General" kg"', '0.000000000000000000000000000000', 'm/d/yyyy', 'h:mm:ss', '# ?/?', '@', '00.00', '#.##',
  ]
  for (const f of codes) for (const b of [UP, DOWN]) one('code', 1.5, f, [b])
  // a selection of cells of other formats: the active cell's result for all
  const mixed = { B2: { value: 1.5, numberFormat: '0.00' }, B3: { value: 2 }, B4: { value: 0.25, numberFormat: '0.0%' } }
  press('selection', mixed, 'B2:B4', [UP])
  press('selection', { ...mixed, B2: { value: 2 } }, 'B2:B4', [DOWN])
  // Comma Style and Accounting
  for (const b of ['Comma Style', 'Accounting Number Format'] as const) {
    one('money', 1234.5, undefined, [b])
    one('money', 1234.5, '0.000', [b])
    one('money', 'abc', undefined, [b])
  }
  one('money', 1234.5, undefined, ['Comma Style', UP])
  // AutoSum on one cell, Return after it
  const sum = (group: string, cells: Record<string, CellInput>, at: string) => add(group, { cells, act: { do: 'press', range: at, buttons: ['AutoSum'], enter: true } })
  const n = (value: number): CellInput => ({ value })
  sum('sum-above', { A1: { value: 'Amount' }, A2: n(1), A3: n(2), A4: n(3) }, 'A5')
  sum('sum-above', { A1: n(1), A2: n(2), A3: n(3), A4: n(4) }, 'A5')
  sum('sum-gap', { A1: n(1), A3: n(3), A4: n(4) }, 'A5')
  sum('sum-gap', { A1: n(1), A2: n(2), A3: n(3) }, 'A5')
  sum('sum-left', { A2: n(1), B2: n(2), C2: n(3) }, 'D2')
  sum('sum-left', { A2: { value: 'Q1' }, B2: n(1), C2: n(2), D2: n(3) }, 'E2')
  sum('sum-both', { D1: n(5), A2: n(1), B2: n(2), C2: n(3) }, 'D2')
  sum('sum-none', { A1: n(1) }, 'H9')
  sum('sum-text', { A1: n(1), A2: n(2), A3: { value: 'x' } }, 'A4')
  sum('sum-subtotals', { A1: n(1), A2: n(2), A3: { formula: 'SUM(A1:A2)' }, A4: n(4), A5: n(5), A6: { formula: 'SUM(A4:A5)' } }, 'A7')
  sum('sum-formulas', { A1: n(1), A2: { formula: 'A1*2' }, A3: { formula: 'A2*2' } }, 'A4')
  sum('sum-dates', { A1: { value: 45688, numberFormat: 14 }, A2: { value: 45689, numberFormat: 14 } }, 'A3')
  return out
}
