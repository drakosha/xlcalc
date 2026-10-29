/**
 * The format a cell takes for what is typed into it, as Excel sets it (the entry sets `typed`, `test/excel/entry/`). A
 * cell in General takes the format of the typed form. A cell whose format is one of Excel's built-in date, time, percent,
 * money, scientific or fraction formats takes the typed form's format when that form is a date, a time, percent, money,
 * scientific or a fraction of another kind (a time typed into a date cell shows `h:mm`), an accounting format by any of
 * them; a plain or grouped number keeps it. The built-in number formats (`0.00`, `#,##0`, `#,##0_);(#,##0)`) and a
 * format of the book's own are kept whatever is typed. A code is a built-in format where Excel takes a file's code so.
 */
import type { NumberFormat } from '../contract.js'
import type { Locale } from '../locale/types.js'
import { BUILTIN_MONEY, builtinIn, EN_BUILTIN, enCode, EN_CODES, isBuiltin } from '../numfmt.js'
import { parseFormat, type Tok } from '../numfmt/parse.js'
import type { Form } from './read.js'

/** The kinds of built-in formats a typed form replaces when it is of another kind; accounting formats are replaced by every form. */
type Family = 'date' | 'time' | 'percent' | 'currency' | 'scientific' | 'fraction' | 'accounting'

const span = (lo: number, hi: number, f: Family): [number, Family][] => Array.from({ length: hi - lo + 1 }, (_, i) => [lo + i, f])

/** The built-in formats by kind; the others (General, the numbers 1 to 4 and 37 to 40, text) are not replaced. */
const FAMILY: ReadonlyMap<number, Family> = new Map([
  ...BUILTIN_MONEY,
  ...span(9, 10, 'percent'),
  [11, 'scientific'],
  [48, 'scientific'],
  ...span(12, 13, 'fraction'),
  ...span(14, 17, 'date'),
  [22, 'date'],
  ...span(27, 31, 'date'),
  [36, 'date'],
  ...span(18, 21, 'time'),
  ...span(32, 35, 'time'),
  ...span(45, 47, 'time'),
])

/** The kind of a typed form that sets a format; a plain or grouped number keeps the cell's format. */
const familyOf = (form: Form): Family | undefined => (form === 'plain' || form === 'grouped' ? undefined : form)

/**
 * The number of the built-in format a cell's format is in a locale: its number, or the first whose code it is in en-US
 * and in the locale alike (`h:mm` is 20 in both, `m/d/yyyy` is 14 in en-US only, as Excel ru-RU keeps `dd.mm.yyyy` a
 * format of the book's own); undefined for a format of the book's own.
 */
export function builtinOf(format: NumberFormat, locale: Locale): number | undefined {
  if (typeof format === 'number') return format
  if (format.toLowerCase() === 'general') return 0
  const code = builtinIn(locale)
  for (let id = 0; isBuiltin(id); id++) if (!BUILTIN_MONEY.has(id) && EN_BUILTIN[id] === format && code(id) === format) return id
  return undefined
}

/** The form of a constant typed with a format of a kind a typed form sets (a formula's format, `formula.ts`); plain for another, which keeps a cell's format. */
export function formOf(format: NumberFormat, locale: Locale): Form {
  const id = builtinOf(format, locale)
  const family = id === undefined ? undefined : FAMILY.get(id)
  return family === undefined ? 'plain' : family === 'accounting' ? 'currency' : family
}

/** The format of a cell with the format `cell` in a locale after a number of the form typed, with that form's format; undefined: General. */
export function formatAfter(form: Form, typed: NumberFormat | undefined, cell: NumberFormat | undefined, locale: Locale): NumberFormat | undefined {
  const id = cell === undefined ? 0 : builtinOf(cell, locale)
  if (id === 0) return typed
  const family = familyOf(form)
  const own = id === undefined ? undefined : FAMILY.get(id)
  return family && own && family !== own ? typed : cell
}

/** Whether a format shows text (`@`), percent or a day (a day, a year or a month's name: a date), by its first section. */
export function shows(format: NumberFormat | undefined, what: 'text' | 'pct' | 'day'): boolean {
  if (format === undefined) return false
  const code = enCode(format)
  const toks: Tok[] = parseFormat(code, EN_CODES)?.[0]?.toks ?? []
  return toks.some((t) => (what === 'day' ? t.t === 'date' && /^(d|y|m{3})/i.test(t.s) : t.t === what))
}
