/**
 * Excel's Increase Decimal and Decrease Decimal (Home > Number): the number format a cell gets, from its format and
 * what it shows. A General cell takes the decimals its number shows in its column, one more or one less (`0.00` for
 * 1.5); a code gets a decimal more or less in each of its sections of digits, after its last digit; a section of a
 * date, a time, a fraction or text stays as it is.
 */
import type { CellValue, NumberFormat } from '../contract.js'
import type { Locale } from '../locale/types.js'
import { columnPoints } from '../column-width.js'
import { builtinIn } from './builtin.js'
import { generalFit } from './general.js'
import { room } from './glyphs.js'
import { splitSections } from './parse.js'

/** The most decimals Excel gives a code. */
const MAX_DECIMALS = 30

/** The number format after one step of Increase Decimal (`step` 1) or Decrease Decimal (-1), in a column of `width` (Excel's standard without one). */
export function stepDecimals(format: NumberFormat | undefined, value: CellValue, step: 1 | -1, locale?: string | Locale, width?: number): NumberFormat {
  const code = format === undefined ? 'General' : typeof format === 'number' ? builtinIn(locale)(format) : format
  if (/^general$/i.test(code)) return fromGeneral(value, step, width)
  const out = splitSections(code).map((s) => stepSection(s, step))
  return out.some((s) => s === null) ? (format ?? 'General') : out.join(';')
}

/** General: the decimals the number shows in its column, one more or one less; a value that is no number shows none. */
function fromGeneral(value: CellValue, step: 1 | -1, width?: number): string {
  const shown = typeof value === 'number' ? generalFit(value, room(columnPoints(width))) : '0'
  const m = /^-?\d*(?:\.(\d*))?(E[+-]\d+)?$/.exec(shown)
  const decimals = Math.max(0, (m?.[1]?.length ?? 0) + step)
  return `0${decimals ? `.${'0'.repeat(decimals)}` : ''}${m?.[2] ? 'E+00' : ''}`
}

/**
 * A section with a decimal more or less, as it is when it has no digits or shows a date, a time or a fraction; null
 * when it has the most decimals already, which leaves the whole code as it is.
 */
function stepSection(section: string, step: 1 | -1): string | null {
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
  if (!digits.length || own.some((i) => /[dmyhsDMYHSbBgGeE/@]/.test(at(i)) && !isExponent(section, i))) return section
  const exp = own.find((i) => isExponent(section, i)) ?? section.length
  const mantissa = digits.filter((i) => i < exp)
  const point = own.find((i) => at(i) === '.' && i < exp)
  // a section of `?` alone after a literal (accounting's zero, `"-"??`) keeps the room of the decimals in them
  const spare = point === undefined && mantissa.every((i) => at(i) === '?')
  const decimals = spare ? mantissa : point === undefined ? [] : mantissa.filter((i) => i > point)
  if (step > 0) {
    if (decimals.length >= MAX_DECIMALS) return null
    const after = decimals.length ? decimals[decimals.length - 1] + 1 : point !== undefined ? point + 1 : mantissa[mantissa.length - 1] + 1
    return `${section.slice(0, after)}${spare ? '?' : point === undefined ? '.0' : '0'}${section.slice(after)}`
  }
  if (spare) return section.slice(0, decimals[decimals.length - 1]) + section.slice(decimals[decimals.length - 1] + 1)
  if (!decimals.length) return point === undefined ? section : section.slice(0, point) + section.slice(point + 1)
  const last = decimals[decimals.length - 1]
  // the point goes with the last decimal
  return decimals.length === 1 ? section.slice(0, point) + section.slice(point! + 1, last) + section.slice(last + 1) : section.slice(0, last) + section.slice(last + 1)
}

/** Whether the `E` or `e` at `i` begins an exponent (`E+`, `E-`). */
const isExponent = (s: string, i: number) => /[Ee]/.test(s[i]) && /[+-]/.test(s[i + 1] ?? '')
