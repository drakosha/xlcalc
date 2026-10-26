/**
 * Excel's Increase Decimal and Decrease Decimal (Home > Number): the number format a cell gets, from its format and
 * what it shows (`test/excel/ui/press-held.json`, `press-more-held.json`). A General cell of a number takes the decimals
 * it shows in its column one more or one less (`0.00` for 1.5); the number's own decimals when the column shows a
 * fraction as a whole number (`12345679` for 12345678.9), scientific notation when it shows `#`; with none left, or
 * no number, it stays. A code gets a decimal more or less in each section of digits after its last one, a
 * copy of it (`#.##` to `#.###`); a section of `?` alone (accounting's zero, `"-"??`), of a date, a time, a fraction
 * or text stays as it is.
 */
import type { CellValue, NumberFormat } from '../contract.js'
import type { Locale } from '../locale/types.js'
import { columnPoints } from '../column-width.js'
import { builtinIn } from './builtin.js'
import { generalFit, generalFormat } from './general.js'
import { room } from './glyphs.js'
import { splitSections } from './parse.js'

/** The number format after one step of Increase Decimal (`step` 1) or Decrease Decimal (-1), in a column of `width` (Excel's standard without one); the format given when the step changes nothing. */
export function stepDecimals(format: NumberFormat | undefined, value: CellValue, step: 1 | -1, locale?: string | Locale, width?: number): NumberFormat | undefined {
  const code = format === undefined ? 'General' : typeof format === 'number' ? builtinIn(locale)(format) : format
  if (/^general$/i.test(code)) return typeof value === 'number' ? (fromGeneral(value, step, width) ?? format) : format
  const out = splitSections(code)
    .map((s) => stepSection(s, step))
    .join(';')
  return out === code ? format : out
}

/** General: the number's decimals one more or one less, null when it has none to take away. */
function fromGeneral(value: number, step: 1 | -1, width?: number): string | null {
  const fit = generalFit(value, room(columnPoints(width)))
  // a fraction the column shows whole: the number's own decimals, as General writes it; `#` across the cell: scientific
  const shown = fit.startsWith('#') ? '0E+00' : /^-?\d+$/.test(fit) && !Number.isInteger(value) ? generalFormat(value) : fit
  const m = /^-?\d*(?:\.(\d*))?(E[+-]\d+)?$/.exec(shown)
  const decimals = (m?.[1]?.length ?? 0) + step
  if (decimals < 0) return null
  return `0${decimals ? `.${'0'.repeat(decimals)}` : ''}${m?.[2] ? 'E+00' : ''}`
}

/** A section with a decimal more or less, as it is when it has no digits but `?` or shows a date, a time or a fraction. */
function stepSection(section: string, step: 1 | -1): string {
  // the places of the section's own characters: outside quotes, escapes, brackets and the characters of `_x` and `*x`
  const own: number[] = []
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') i = section.indexOf('"', i + 1) < 0 ? section.length : section.indexOf('"', i + 1)
    else if (ch === '[') i = section.indexOf(']', i) < 0 ? section.length : section.indexOf(']', i)
    else if (ch === '\\' || ch === '_' || ch === '*') i++
    else own.push(i)
  }
  const at = (i: number) => section[i]
  const digits = own.filter((i) => '0#?'.includes(at(i)))
  if (!digits.some((i) => at(i) !== '?') || own.some((i) => /[dmyhsDMYHSbBgGeE/@]/.test(at(i)) && !isExponent(section, i))) return section
  const exp = own.find((i) => isExponent(section, i)) ?? section.length
  const mantissa = digits.filter((i) => i < exp)
  const point = own.find((i) => at(i) === '.' && i < exp)
  const decimals = point === undefined ? [] : mantissa.filter((i) => i > point)
  const last = decimals[decimals.length - 1]
  if (step > 0) {
    if (point === undefined) return `${section.slice(0, mantissa[mantissa.length - 1] + 1)}.0${section.slice(mantissa[mantissa.length - 1] + 1)}`
    const after = last === undefined ? point + 1 : last + 1
    return `${section.slice(0, after)}${last === undefined ? '0' : at(last)}${section.slice(after)}`
  }
  if (last === undefined) return point === undefined ? section : section.slice(0, point) + section.slice(point + 1)
  // the point goes with the last decimal
  return decimals.length === 1 ? section.slice(0, point) + section.slice(point! + 1, last) + section.slice(last + 1) : section.slice(0, last) + section.slice(last + 1)
}

/** Whether the `E` or `e` at `i` begins an exponent (`E+`, `E-`). */
const isExponent = (s: string, i: number) => /[Ee]/.test(s[i]) && /[+-]/.test(s[i + 1] ?? '')
