/**
 * The format a cell takes for what is typed into it, as Excel sets it: a cell in General takes the format of the typed
 * form (a date, a time, percent, money, grouping, scientific, a fraction); a cell with a format keeps it, unless a date
 * or a time comes into a cell that shows no date or time, or percent or money into one that shows none. A plain number
 * typed into a percent cell is taken as percent when it is 1 or more (Excel's automatic percent entry: 10 is 10%).
 */
import type { NumberFormat } from '../contract.js'
import { EN_BUILTIN, EN_CODES } from '../numfmt.js'
import { parseFormat, type Tok } from '../numfmt/parse.js'
import type { Form, Typed } from './read.js'

/** What a cell's format shows, by its first section: General, text, a date, a time, percent, money, scientific, a fraction, a number. */
export type Category = 'general' | 'text' | Exclude<Form, 'plain' | 'grouped'> | 'number'

/** Whether a section's date parts show a date: a year, a day, or a month (`m` is minutes after hours or before seconds). */
function showsDate(toks: Tok[]): boolean {
  const parts = toks.flatMap((t) => (t.t === 'date' || t.t === 'elapsed' ? [t.s[0].toLowerCase()] : []))
  return parts.some((c, i) => c === 'y' || c === 'd' || (c === 'm' && parts[i - 1] !== 'h' && parts[i + 1] !== 's'))
}

/** The category of a cell's format (an en-US code as xlsx stores it, or a built-in format's number). */
export function category(format: NumberFormat | undefined): Category {
  if (format === undefined) return 'general'
  const code = typeof format === 'number' ? EN_BUILTIN[format] : format
  const toks: Tok[] = parseFormat(code, EN_CODES)?.[0]?.toks ?? []
  if (toks.some((t) => t.t === 'text')) return 'text'
  if (toks.some((t) => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm')) return showsDate(toks) ? 'date' : 'time'
  if (toks.some((t) => t.t === 'general') || !toks.some((t) => t.t === 'ph')) return 'general'
  if (toks.some((t) => t.t === 'pct')) return 'percent'
  if (toks.some((t) => t.t === 'exp')) return 'scientific'
  if (toks.some((t) => t.t === 'lit' && /[$€£¥₽]/.test(t.s)) || /\[\$[^-\]]/.test(code)) return 'currency'
  if (toks.some((t) => t.t === 'slash')) return 'fraction'
  return 'number'
}

/** The forms that replace a cell's format of another category. */
const REPLACING: ReadonlySet<Form> = new Set(['date', 'time', 'percent', 'currency'])

const dated = (c: Category) => c === 'date' || c === 'time'

/** The format of a cell with the format `cell` after a number of the form typed; undefined: General. */
export function formatAfter(typed: Extract<Typed, { t: 'number' }>, cell: NumberFormat | undefined): NumberFormat | undefined {
  const c = category(cell)
  if (c === 'general') return typed.format ?? cell
  if (!REPLACING.has(typed.form)) return cell
  const same = dated(c) ? typed.form === 'date' || typed.form === 'time' : c === typed.form
  return same ? cell : typed.format
}

/** A plain number typed into a percent cell: from 1 up it is taken as percent (10 is 10%), below 1 as it is (0.1 is 10% too). */
export function percentEntry(typed: Extract<Typed, { t: 'number' }>, cell: NumberFormat | undefined): number {
  return category(cell) === 'percent' && (typed.form === 'plain' || typed.form === 'grouped') && Math.abs(typed.v) >= 1 ? typed.v / 100 : typed.v
}
