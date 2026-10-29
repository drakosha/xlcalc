/**
 * Held out from rotation-held, after the rules its misses gave (a margin's projection of a half rounds up; a number
 * shown by a format other than a date's takes two points more, as one in General does), predicted before Excel ran
 * it: sizes whose margin projects to a half at 30 degrees, up and down, and numbers by formats of every kind turned.
 */
import { collector, one, type RowCase, size } from './cases.js'
import { turned } from './rotation.js'

const FORMATS = ['0.00', '0%', '#,##0', '$#,##0.00', '0.00E+00', '# ?/?', 'h:mm', 'mm/dd/yyyy', '@', '"x"0', '0;-0;"zero"']

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  for (const n of [18, 20, 24, 34, 38]) for (const a of [30, -30]) add('half', { cells: one('bbbb', turned(a, size(n))) })
  for (const f of FORMATS) for (const a of [20, 45, 70]) add('format', { cells: { 0: { value: 1234.5, format: f, look: turned(a) } } })
  add('format', { cells: { 0: { value: 0, format: '0;-0;"zero"', look: turned(45) } } })
  return out
}
