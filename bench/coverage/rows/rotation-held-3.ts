/**
 * Held out from rotation-held-2, after the rule its miss gave (a number shown by a section of text alone takes no two
 * points more), predicted before Excel ran it: sections of literal text, of text with digits and of a time turned.
 */
import { collector, type RowCase } from './cases.js'
import { turned } from './rotation.js'

const CASES: [number, string][] = [
  [0, '0;-0;"none"'],
  [-3, '0;"minus"'],
  [5, '"five"'],
  [5, '"n="0'],
  [0.5, '[h]:mm'],
  [12, '0 "kg"'],
  [-3, '0;(0)'],
]

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  for (const [value, format] of CASES) for (const a of [30, 60]) add('format', { cells: { 0: { value, format, look: turned(a) } } })
  return out
}
