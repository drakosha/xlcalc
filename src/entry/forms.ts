/**
 * The format a cell takes for what is typed into it, as Excel sets it (the entry sets `typed`, `test/excel/entry/`). A
 * cell in General takes the format of the typed form. A cell whose format is one of Excel's built-in date, time, percent,
 * money, scientific or fraction formats takes the typed form's format when that form is a date, a time, percent, money,
 * scientific or a fraction of another kind (a time typed into a date cell shows `h:mm`), an accounting format by any of
 * them; a plain or grouped number keeps it. The built-in number formats (`0.00`, `#,##0`, `#,##0_);(#,##0)`) and a
 * format of the book's own are kept whatever is typed. A code is a built-in format where Excel takes a file's code so.
 */
import type { NumberFormat } from '../contract.js'
import { REGIONAL } from '../locale/rules.js'
import { builtinCodes, EN_BUILTIN, EN_CODES, isBuiltin } from '../numfmt.js'
import { parseFormat, type Tok } from '../numfmt/parse.js'
import type { Form } from './read.js'

/** The kinds of built-in formats a typed form replaces when it is of another kind; accounting formats are replaced by every form. */
type Family = 'date' | 'time' | 'percent' | 'currency' | 'scientific' | 'fraction' | 'accounting'

const span = (lo: number, hi: number, f: Family): [number, Family][] => Array.from({ length: hi - lo + 1 }, (_, i) => [lo + i, f])

/** The built-in formats by kind; the others (General, the numbers 1 to 4 and 37 to 40, text) are not replaced. */
const FAMILY: ReadonlyMap<number, Family> = new Map([
  ...span(5, 8, 'currency'),
  ...span(41, 44, 'accounting'),
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

/** The kind of each typed form that sets a format. */
const FORM: Partial<Record<Form, Family>> = { date: 'date', time: 'time', percent: 'percent', currency: 'currency', scientific: 'scientific', fraction: 'fraction' }

/**
 * The number of the built-in format a cell's format is in the locale `tag`: its number, or the first whose code it is
 * in en-US and in the locale alike (`h:mm` is 20 in both, `m/d/yyyy` is 14 in en-US only, as Excel ru-RU keeps
 * `dd.mm.yyyy` a format of the book's own); undefined for a format of the book's own.
 */
export function builtinOf(format: NumberFormat, tag: string): number | undefined {
  if (typeof format === 'number') return format
  if (format.toLowerCase() === 'general') return 0
  const own = REGIONAL.get(tag)?.builtin
  const code = own ? builtinCodes(own) : (id: number) => EN_BUILTIN[id]
  for (let id = 0; isBuiltin(id); id++) if (EN_BUILTIN[id] === format && code(id) === format) return id
  return undefined
}

/** The format of a cell with the format `cell` in the locale `tag` after a number of the form typed, with that form's format; undefined: General. */
export function formatAfter(form: Form, typed: NumberFormat | undefined, cell: NumberFormat | undefined, tag: string): NumberFormat | undefined {
  const id = cell === undefined ? 0 : builtinOf(cell, tag)
  if (id === 0) return typed
  const family = FORM[form]
  const own = id === undefined ? undefined : FAMILY.get(id)
  return family && own && family !== own ? typed : cell
}

/** Whether a format shows text (`@`) or percent, by its first section. */
export function shows(format: NumberFormat | undefined, what: 'text' | 'pct'): boolean {
  if (format === undefined) return false
  const code = typeof format === 'number' ? EN_BUILTIN[format] : format
  const toks: Tok[] = parseFormat(code, EN_CODES)?.[0]?.toks ?? []
  return toks.some((t) => t.t === what)
}
