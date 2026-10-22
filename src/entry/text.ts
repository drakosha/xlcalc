/**
 * The text a cell opens with for editing (`entryText`), as Excel's formula bar shows a value: a number whole (15
 * significant digits, positional up to 21 characters), a percent as a percent, a date or a time by the locale's short
 * date and long time (two spaces between them, as the bar writes them), booleans and errors in the locale's words, and
 * text with an apostrophe where typing it would read otherwise (`needsApostrophe`, by `parseEntry`). The value reads
 * back from the text edited in place (`parseEntry` with `inPlace`).
 */
import type { CellValue, EntryTextOptions, NumberFormat } from '../contract.js'
import { epochOf } from '../dates.js'
import { checkCellValue, checkEntryTextOptions } from '../input.js'
import { checkString } from '../input/basic.js'
import { DEFAULT_LOCALE } from '../locale/registry.js'
import { displayRules, type Rules } from '../locale/rules.js'
import { EN_BUILTIN, EN_CODES, generalText } from '../numfmt.js'
import { minutes } from '../numfmt/date.js'
import { parseFormat, type Tok } from '../numfmt/parse.js'
import { choose, kind } from '../numfmt/select.js'
import { parseEntry } from './index.js'

/** The first serial past 9999-12-31, where Excel's calendar ends. */
const END = 2_958_466

/** What a number's section of its format shows it as: a date (a day, maybe a time), a time alone, a percent, a number. */
function shownAs(n: number, format: NumberFormat | undefined): 'date' | 'time' | 'percent' | 'number' {
  if (format === undefined) return 'number'
  const sec = choose(parseFormat(typeof format === 'number' ? EN_BUILTIN[format] : format, EN_CODES) ?? [], n)?.sec
  if (!sec) return 'number'
  if (kind(sec.toks) === 'date') return sec.toks.some((t, i) => t.t === 'date' && (t.s[0] === 'y' || t.s[0] === 'd' || (t.s[0] === 'm' && !minutes(sec.toks, i)))) ? 'date' : 'time'
  return sec.toks.some((t: Tok) => t.t === 'pct') ? 'percent' : 'number'
}

/** A number as the bar writes it: positional up to 21 characters, the locale's decimal separator. */
const whole = (n: number, r: Rules) => generalText(n, 21).replace('.', r.decimal)

/** A date and time as the formula bar shows it: the short date (built-in format 14) and, past midnight, two spaces and the long time. */
function dateTime(n: number, time: boolean, r: Rules): string {
  // the fraction rounded half up, not `momentOf`: the formula bar was not measured at a second's edge (the set `seconds-held`)
  let day = Math.floor(n)
  let secs = Math.round((n - day) * 86_400)
  if (secs === 86_400) [day, secs] = [day + 1, 0]
  // rounded past 9999-12-31, the value is past the calendar: the number, as for any value there
  if (day + r.epoch >= END) return whole(n, r)
  const clock = r.display(secs / 86_400, '[$-F400]h:mm:ss AM/PM', Infinity).text
  if (time && day === 0) return clock
  const date = r.display(day, 14, Infinity).text
  return secs || time ? `${date}  ${clock}` : date
}

function numberText(n: number, format: NumberFormat | undefined, r: Rules): string {
  const as = shownAs(n, format)
  if (as === 'percent') return `${whole(n * 100, r)}%`
  if ((as === 'date' || as === 'time') && n >= 0 && n + r.epoch < END) return dateTime(n, as === 'time', r)
  return whole(n, r)
}

/** Whether text typed into a cell with these options would read as something else than that text (`parseEntry`). */
function readsOtherwise(text: string, o: EntryTextOptions): boolean {
  const r = parseEntry(text, { ...(o.locale ? { locale: o.locale } : {}), ...(o.dateSystem ? { dateSystem: o.dateSystem } : {}), ...(o.now ? { now: o.now } : {}), ...(o.numberFormat === undefined ? {} : { cell: { numberFormat: o.numberFormat } }) })
  return !(r.ok && r.input?.value === text && r.input.formula === undefined)
}

/**
 * Whether text typed into a cell would be parsed as something other than that text, so the cell needs an apostrophe
 * before it to keep it as text: a number, a date, a boolean, an error or a formula in the locale (`007`, `1.5` in
 * ru-RU, `=A1`), a number in a percentage cell, text that begins with an apostrophe, or empty text.
 */
export function needsApostrophe(text: string, opts?: EntryTextOptions): boolean {
  return readsOtherwise(checkString(text, 'text'), checkEntryTextOptions(opts))
}

/**
 * The text a cell holding `value` opens with in edit mode, as Excel's formula bar displays it, in `opts.locale` and
 * according to the cell's `opts.numberFormat`: a number to 15 significant digits with the locale's decimal separator
 * (`1234.5`, `1234,5`), a percentage cell's number as a percentage (`12%`), a date or time cell's value in the locale's
 * short date and long time formats (`1/31/2025  6:00:00 PM`, `31.01.2025  18:00:00`), a boolean or an error in the
 * locale's words, and text with an apostrophe where typing it would otherwise be parsed as something else (`'007`).
 * Parsing the text as edited in place gives back the value (`parseEntry` with `inPlace: true`), but for a negative
 * percentage, which Excel's in-place editor reads as a formula (`-5%` is `=-5%%`).
 */
export function entryText(value: CellValue, opts?: EntryTextOptions): string {
  return checkedEntryText(checkCellValue(value, 'value'), checkEntryTextOptions(opts))
}

/** `entryText` of a value and options already checked (a window's cells, `ViewportCell.editText`). */
export function checkedEntryText(v: CellValue, o: EntryTextOptions): string {
  if (v === null) return ''
  if (typeof v === 'string') return readsOtherwise(v, o) ? `'${v}` : v
  const r = displayRules(o.locale ?? DEFAULT_LOCALE, epochOf(o.dateSystem))
  if (typeof v === 'boolean') return v ? r.locale.booleans.TRUE : r.locale.booleans.FALSE
  if (typeof v === 'object') return r.errorText(v.error)
  return numberText(v, o.numberFormat, r)
}
