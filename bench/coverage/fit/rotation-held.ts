/**
 * Held out from the fit sets rotation and rotation-2, predicted by the rules they gave before Excel ran it: angles and
 * text widths not measured, sizes between those measured, turned text of other scripts, numbers by formats and booleans
 * turned, wrapped text turned up or down over line breaks, several turned cells and a level one, a turned cell in a
 * hidden row, and numbers typed into turned cells of other angles.
 */
import type { CellInput } from '../../../src/contract.js'
import { turned } from '../rows/rotation.js'
import { textOf } from './autofit.js'
import { collector, HIDDEN, type FitCase, type FitCell } from './cases.js'

const one = (value: CellInput['value'], look: FitCell['look'], more: Partial<FitCell> = {}): Record<number, FitCell> => ({ 1: { value, look, ...more } })

export function buildFitCases(): FitCase[] {
  const { out, add } = collector()
  for (const a of [7, 12, 14, 16, 33, 47, 49, 52, 77, -7, -33, -52]) for (const w of [23, 77]) add('angle', { cells: one(textOf(w), turned(a)) })
  for (const a of [30, 60]) for (const w of [13, 27, 61]) add('half', { cells: one(textOf(w), turned(a)) })
  for (const n of [11.5, 17, 22, 26, 30, 40]) for (const a of [30, 75, 90, 'vertical'] as const) add('size', { cells: one('bbbbbbbb', turned(a, { font: { size: n } })) })
  for (const t of ['Итого', 'Сумма к оплате']) for (const a of [45, 90] as const) add('script', { cells: one(t, turned(a)) })
  for (const a of [30, 90] as const) {
    add('format', { cells: one(0.1234, turned(a), { format: '0.00%' }) })
    add('format', { cells: one(1234.5, turned(a), { format: '#,##0.00' }) })
    add('format', { cells: one(false, turned(a)) })
    add('format', { cells: one(-1234567, turned(a, { font: { size: 16 } })) })
  }
  for (const a of [90, -90] as const) add('wrap', { cells: one('a\nb\nc', turned(a, { alignment: { wrap: true } })) })
  add('wrap', { cells: one('one\ntwo', turned(90, { alignment: { wrap: true }, font: { size: 16 } })) })
  add('cells', { cells: { 1: { value: textOf(60), look: turned(30) }, 2: { value: textOf(80), look: turned(60) }, 3: { value: textOf(40) } } })
  add('cells', { cells: { 1: { value: textOf(200), look: turned(75) }, 2: { value: 'bbb', look: turned('vertical', { font: { size: 20 } }) } } })
  add('hidden', { cells: { [HIDDEN]: { value: textOf(100), look: turned(20) } } })
  for (const a of [30, 'vertical'] as const) add('typed', { cells: { 1: { look: turned(a) } }, typed: [[1, '123456789012']], via: 'value' })
  return out
}
