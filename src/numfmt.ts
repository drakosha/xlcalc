/**
 * Number formats (the category's assembly point; the pieces are in `numfmt/`): TEXT's `formatWith`, by a
 * dialect's codes, and a cell's display `displayWith`, by an xlsx file's codes in the book's locale, with the
 * width of its column, `#` for a number that does not fit, fills and the section's colour. One parser and one
 * formatter serve both; the rules and what was measured in Excel are in `docs/formats.md`.
 */
import type { CellDisplay, CellValue, NumberFormat } from './contract.js'
import { formatDate } from './numfmt/date.js'
import { generalFit } from './numfmt/general.js'
import { DIGIT, fitText, glyph, hashes, MAX_SHOWN, room, textWidth } from './numfmt/glyphs.js'
import { formatNumber } from './numfmt/number.js'
import { EN_CODES, parseFormat, type Codes, type Section } from './numfmt/parse.js'
import { choose, kind, textSection } from './numfmt/select.js'
import { FILL, plain, type Style } from './numfmt/style.js'

export { builtinCodes, EN_BUILTIN, isBuiltin } from './numfmt/builtin.js'
export { generalFormat, generalText } from './numfmt/general.js'
export { columnPoints, room, widthOf } from './numfmt/glyphs.js'
export { EN_CODES, type Codes } from './numfmt/parse.js'
export { EN_NAMES, styleIn, type Names, type Style } from './numfmt/style.js'

/** How TEXT reads and writes format codes in a locale (a cell reads an xlsx file's codes, `EN_CODES`). */
export interface Dialect {
  codes: Codes
  style: Style
  /** Text read as a number (TEXT's text argument), or null. */
  number(text: string): number | null
}

/** A section with the locale's long date or time in place of `[$-F800]` and `[$-F400]`, and its month names. */
function dated(sec: Section, style: Style) {
  const long = sec.lcid === 'F800' ? style.longDate : sec.lcid === 'F400' ? style.longTime : undefined
  const names = (sec.lcid && style.names(sec.lcid)) || style
  if (long === undefined) return { sec, names }
  return { sec: parseFormat(long, EN_CODES)![0], names: sec.lcid === 'F800' && style.longMonths ? { ...names, months: style.longMonths } : names }
}

/** Text by a section (`@` the text itself); its fills still marked. */
const asText = (sec: Section, text: string, style: Style) => sec.toks.map((t) => (t.t === 'text' ? text : plain(t, style))).join('')

/** Fills dropped, as TEXT writes them. */
const unfilled = (text: string) => text.replace(/\uFFFF[^]?/g, '')

/**
 * A number by the section its value chooses: the text with fills marked, or General (the section's minus aside,
 * `general`), or undefined when the section cannot show it; the section gives the colour.
 */
function byChoice(v: number, secs: Section[], style: Style): { text?: string; general?: number; sec: Section | null } {
  const choice = choose(secs, v)
  if (!choice) return { sec: null }
  const { sec, signed } = choice
  const general = signed ? v : Math.abs(v)
  if (!sec) return { general, sec }
  const k = kind(sec.toks)
  if (k === 'general' && sec.toks.length === 1) return { general, sec }
  // a number under a text section shows as General, without the section's colour
  if (k === 'text') return { general, sec: null }
  const sign = signed && v < 0 ? '-' : ''
  if (k === 'general') return { text: sign + sec.toks.map((t) => (t.t === 'general' ? style.general(Math.abs(v)) : plain(t, style))).join(''), sec }
  if (k === 'date') {
    const d = dated(sec, style)
    return { text: formatDate(v, d.sec, style, d.names) ?? undefined, sec }
  }
  if (k === 'literal') return { text: sign + sec.toks.map((t) => plain(t, style)).join(''), sec }
  return { text: formatNumber(v, sec.toks, signed, style) ?? undefined, sec }
}

/** A number by its format's sections (`byChoice`), General without them. */
const byFormat = (v: number, secs: Section[] | null, style: Style): ReturnType<typeof byChoice> => (secs ? byChoice(v, secs, style) : { general: v, sec: null })

/** TEXT(value, format) in a dialect; null for a format it does not read or a value the format cannot show (#VALUE!). */
export function formatWith(value: number | string | boolean, fmt: string, d: Dialect): string | null {
  const secs = parseFormat(fmt, d.codes)
  if (!secs) return null
  const { style } = d
  if (typeof value === 'boolean') return style.bool(value)
  if (fmt === '') return ''
  if (typeof value === 'string') {
    const n = d.number(value)
    if (n === null) {
      const sec = textSection(secs)
      return sec ? unfilled(asText(sec, value, style)) : value
    }
    value = n
  }
  const { text, general } = byChoice(value, secs, style)
  if (general !== undefined) return style.general(general)
  if (text === undefined) return null
  const out = unfilled(text)
  return out.length > MAX_SHOWN ? null : out
}

/**
 * A fill repeated across the room the rest leaves, the last fill only (earlier ones drop out); null when the rest
 * does not fit, or the text would pass 255 characters.
 */
function filled(text: string, space: number): string | null {
  const at = text.lastIndexOf(FILL)
  if (at < 0) return textWidth(text) <= space && text.length <= MAX_SHOWN ? text : null
  const ch = text[at + 1] ?? ''
  const before = unfilled(text.slice(0, at))
  const after = text.slice(at + 2)
  const free = space - textWidth(before + after)
  if (free < 0) return null
  const out = before + (ch ? ch.repeat(Math.floor(free / glyph(ch))) : '') + after
  return out.length <= MAX_SHOWN ? out : null
}

/** A word (a boolean, an error) shown whole when it fits, else `#` across the cell; text overflows instead. */
const fitted = (text: string, space: number, overflow: boolean) => filled(text, space) ?? (overflow ? unfilled(text) : hashes(space))

/** A cell's format as sections: an xlsx file's code, or a built-in format by its number in the style's locale; null for none (General) and for a code that does not parse. */
const sectionsOf = (format: NumberFormat | undefined, style: Style) => {
  const fmt = typeof format === 'number' ? style.builtin(format) : format
  return fmt === undefined ? null : parseFormat(fmt, EN_CODES)
}

/**
 * What a cell shows: a value by its format (an xlsx file's code, a built-in format by its number, General without one)
 * in `space` points of room, written in the style of the book's locale. Text is never cut (it overflows); a number
 * that does not fit shows `#`.
 */
export function displayWith(value: CellValue, format: NumberFormat | undefined, style: Style, space: number): CellDisplay {
  if (value === null) return { text: '' }
  if (typeof value === 'object') return { text: fitted(style.error(value.error), space, false) }
  const secs = sectionsOf(format, style)
  if (typeof value !== 'number') {
    const sec = secs && textSection(secs)
    // text without a text section shows as it is, whatever the width
    if (!sec && typeof value === 'string') return { text: value }
    const word = typeof value === 'boolean' ? style.bool(value) : value
    const text = fitted(sec ? asText(sec, word, style) : word, space, typeof value === 'string')
    return sec?.color ? { text, color: sec.color } : { text }
  }
  const { text, general, sec } = byFormat(value, secs, style)
  // a value its section cannot show (a date out of range) fills the cell with `#` in the automatic colour
  if (text === undefined && general === undefined) return { text: hashes(space) }
  const shown = general !== undefined ? generalFit(general, space).replace('.', style.decimal) : (filled(text!, space) ?? hashes(space))
  return sec?.color ? { text: shown, color: sec.color } : { text: shown }
}

/** Styles that lay text out for AutoFit, by the style they lay out. */
const LAID = new WeakMap<Style, Style>()

/**
 * The points of the column AutoFit gives a cell (`test/excel/fit/`): its text as `fitText` measures it, a padding `_x` as
 * wide as `x` and a `?` with no digit as a digit; General at its full length, each character a digit and a leading minus
 * by its width, two points past the padding, a scientific form with the padding alone. Null when the cell shows nothing:
 * a blank, empty text, a number its section cannot show (`#` in any width).
 */
export function fitPoints(value: CellValue, format: NumberFormat | undefined, style: Style): number | null {
  if (value === null) return null
  let laid = LAID.get(style)
  if (!laid) LAID.set(style, (laid = { ...style, layout: true }))
  const secs = sectionsOf(format, style)
  let text: string | undefined
  if (typeof value === 'object') text = style.error(value.error)
  else if (typeof value !== 'number') {
    const sec = secs && textSection(secs)
    const word = typeof value === 'boolean' ? style.bool(value) : value
    text = sec ? asText(sec, word, laid) : word
  } else {
    const chosen = byFormat(value, secs, laid)
    if (chosen.general !== undefined) {
      const g = style.general(Math.abs(chosen.general))
      const w = (chosen.general < 0 ? glyph('-') : 0) + DIGIT * g.length
      return g.includes('E') ? w + 5 : w + 7
    }
    text = chosen.text
  }
  return text ? fitText(text) : null
}

/**
 * The points Excel gives a column of `now` points as a number is typed into one of its cells (`test/excel/fit/typed.json`,
 * `held.json`): the points AutoFit gives the cell when they are more, though the cell shows the number whole, but for a
 * number in General the column shows with fewer decimals; null when the column keeps its width.
 */
export function typedPoints(value: CellValue, format: NumberFormat | undefined, style: Style, now: number): number | null {
  if (typeof value !== 'number') return null
  const points = fitPoints(value, format, style)
  if (points === null || points <= now) return null
  const { general } = byFormat(value, sectionsOf(format, style), style)
  if (general === undefined) return points
  const shown = generalFit(general, room(now))
  return shown !== generalFit(general, Infinity) && !/[E#]/.test(shown) ? null : points
}
