/**
 * Text typed into a cell as Excel takes it in a book's locale (`parseEntry`), measured in the entry sets: a formula after
 * `=`, or after a `+` or `-` that begins no number (typed so, Excel makes `+1+2` the formula `1+2` and `-C1` the formula
 * `-C1`), translated from the locale's formula text; text after an apostrophe, or anything in a cell formatted as text;
 * else a constant (`read.ts`), whose form may set the cell's format (`forms.ts`). A number typed into a percent cell
 * gets the `%` Excel's editor adds as it is typed (`5` is 5%). The cell keeps its style, meta, note and link. Only a
 * formula that does not parse stops the entry, as Excel's dialog does; an unknown function or name is taken, and the
 * cell computes `#NAME?`. Text pasted from the clipboard is read so too, but where Excel reads it otherwise (`readEntry`).
 */
import type { CellInput, CheckCode, EntryOptions, EntryResult, ExternalBook, NumberFormat } from '../contract.js'
import { checkEntryOptions, checkString } from '../input.js'
import { validateTyped, toInvariant } from '../locale/index.js'
import { validate } from '../syntax.js'
import { localeOf } from '../locale/registry.js'
import { markingClock, type RulesClock } from '../locale/rules/data.js'
import { rulesOf, type Rules } from '../locale/rules.js'
import type { Locale } from '../locale/types.js'
import { epochOf, ymd } from '../dates.js'
import { FUNCTIONS } from '../functions.js'
import { msToSerial } from '../model.js'
import { formatAfter, shows } from './forms.js'
import { closeTyped } from './close.js'
import { formulaFormat } from './formula.js'
import { moneyFormat, readTyped, type Form } from './read.js'

/** The rules of each locale object and date system, with a clock whose year is set for each entry (entries are synchronous). */
const RULES = new WeakMap<Locale, Map<number, { rules: Rules; clock: RulesClock & { y: number } }>>()

function rulesFor(locale: Locale, epoch: number, year: number) {
  let byEpoch = RULES.get(locale)
  if (!byEpoch) RULES.set(locale, (byEpoch = new Map()))
  let r = byEpoch.get(epoch)
  if (!r) {
    const clock: RulesClock & { y: number } = Object.assign(markingClock(() => clock.y), { y: year })
    byEpoch.set(epoch, (r = { rules: rulesOf(locale, clock, epoch), clock }))
  }
  r.clock.y = year
  return r
}

/** The fields of the cell typed into that its new input keeps. */
const KEPT = ['style', 'meta', 'note', 'link'] as const

/** The check codes of a formula Excel does not take as typed: its dialog stops the entry. */
const STOPS: ReadonlySet<CheckCode> = new Set(['syntax', 'empty', 'too_long', 'too_deep'])

/**
 * A typed formula, `shift` characters into the typed text: its canonical stored text and the format it sets (`formula.ts`),
 * its invariant text where it names what the book may lack (`#NAME?` then), or why Excel stops it, at its place in the text.
 */
function formulaOf(text: string, shift: number, locale: Locale, externalBooks?: ExternalBook[], close = false): Formula | Extract<EntryResult, { ok: false }> {
  const opts = externalBooks ? { externalBooks } : undefined
  const v = validateTyped(text, locale, opts)
  // typed or edited in place, what is left open at the end Excel closes on Return, or offers to (`close.ts`)
  const c = close && !v.ok && v.code === 'syntax' ? closeTyped(text) : null
  if (c?.silent) {
    const f = taken(c.closed, validateTyped(c.closed, locale, opts), shift, locale, externalBooks)
    return 'formula' in f ? f : { ...f, at: Math.min(f.at, text.length + shift) }
  }
  const f = taken(text, v, shift, locale, externalBooks)
  return c?.offered && !('formula' in f) && c.calls.every((n) => known(n, locale)) && validateTyped(c.correction, locale, opts).ok ? { ...f, correction: c.correction } : f
}

/** A typed formula's canonical stored text and the format it sets. */
type Formula = { formula: string; sets?: NumberFormat | 'money' }

/** `text` as `validateTyped` read it (`v`): taken as a formula, or why Excel stops it, at its place in the typed text. */
function taken(text: string, v: ReturnType<typeof validateTyped>, shift: number, locale: Locale, externalBooks?: ExternalBook[]): Formula | Extract<EntryResult, { ok: false }> {
  if (v.ok) {
    const p = validate(v.canonical)
    return { formula: v.canonical, ...(p.ok ? { sets: formulaFormat(p.ast) } : {}) }
  }
  // a book that is not one of its links: Excel would ask for the file, a new link the model does not make
  if (STOPS.has(v.code) || (externalBooks && v.code === 'external_ref')) return { ok: false, code: v.code, message: v.message, at: v.at + shift }
  const t = toInvariant(text, locale)
  return t.ok ? { formula: t.formula } : { ...t, at: t.at + shift }
}

/** Whether `name` is a function's name in the locale, in any case: its own word, else its English one (`SUM` is none in ru-RU, `ABS` is). */
function known(name: string, locale: Locale): boolean {
  const upper = name.toUpperCase()
  return upper in FUNCTIONS ? (locale.functions[upper] ?? upper).toUpperCase() === upper : Object.values(locale.functions).some((w) => w.toUpperCase() === upper)
}

/** Whether typed text begins with a `+` before a number (`+1*2`, `+.5+1`, `++5`): Excel drops it, the formula is `1*2`. */
const plusNumber = (text: string, decimal: string) => text.startsWith('+') && /^\+?[\d.]/.test(text.slice(1).replace(decimal, '.'))

/** The format Excel gives `€5` pasted, in en-US and ru-RU alike (the set `ui/paste-held`). */
const PASTED_EURO = '[$€-2]\\ #,##0;[Red]\\-[$€-2]\\ #,##0'

/**
 * Text typed into a cell as Excel takes it in the book's locale: the cell's new input, with the cell's style, meta, note
 * and link and the format Excel leaves it, or a formula Excel does not take. Empty text clears the value and keeps the rest.
 */
export function parseEntry(text: string, opts?: EntryOptions): EntryResult {
  checkString(text, 'text')
  return readEntry(text, checkEntryOptions(opts))
}

/** How text comes into a cell: typed over it, edited in place (F2, a double click, the formula bar), pasted from the clipboard, put back by Excel's Replace. */
export type EntryWay = 'typed' | 'inPlace' | 'pasted' | 'replaced'

/**
 * Text that comes into a cell (options checked), typed over it, edited in place or pasted from the clipboard (`way`).
 * Edited in place, Excel's percent editor adds `%` only to a number begun by a sign (the set `ui/edit-held`). Pasted, as
 * Excel pastes text (the set `ui/paste-held`, Excel for Mac and Windows): read as typed, but an apostrophe stays part of
 * the text, a formula Excel does not take is text, a percent cell has no percent editor (`5` is 5, 500%), and a euro
 * sign before a whole number (`€5`) gives the euro format Excel gives it; the other forms with a euro sign were not
 * measured pasted and are read as typed. Put back by Excel's Replace (the sets `replace/rules-held` and `rules-held-2`):
 * read as typed, but a percent cell has no percent editor (`6%` is 0.06, `5` 500%, `-5` -500%) and a cell of the text
 * format takes a number, a date, a boolean or a formula as a General cell does, keeping its format where typing gives
 * none (`3/15/2026` is the date with its date format, `12` the number shown by `@`).
 */
export function readEntry(text: string, o: EntryOptions, way: EntryWay = 'typed'): EntryResult {
  const pasted = way === 'pasted'
  const locale = localeOf(o.locale)
  const cell = o.cell ?? {}
  // what goes with the cell, as Excel keeps it whatever is typed
  const kept: CellInput = {}
  for (const k of KEPT) if (cell[k] !== undefined) (kept as Record<string, unknown>)[k] = cell[k]
  const input = (body: CellInput, format: NumberFormat | undefined = cell.numberFormat): EntryResult => ({ ok: true, input: { ...body, ...(format !== undefined ? { numberFormat: format } : {}), ...kept } })
  // put back by Replace, a cell of the text format takes the format typing gives a General cell, and keeps its own where that gives none
  const base = way === 'replaced' && shows(cell.numberFormat, 'text') ? undefined : cell.numberFormat
  const after = (form: Form, typed: NumberFormat | undefined) => formatAfter(form, typed, base, locale) ?? (base === cell.numberFormat ? undefined : cell.numberFormat)
  const epoch = epochOf(o.dateSystem)
  const year = ymd(Math.floor(msToSerial(o.now === undefined ? Date.now() : Date.parse(o.now)))).y
  const { rules, clock } = rulesFor(locale, epoch, year)
  // a formula takes the format of what it computes as a typed constant of that kind would (`forms.ts`)
  const formula = (src: string, shift: number): EntryResult => {
    const f = formulaOf(src, shift, locale, o.externalBooks, way === 'typed' || way === 'inPlace')
    if (!('formula' in f)) return pasted ? input({ value: text }) : f
    if (f.sets === undefined) return input({ formula: f.formula })
    const sets = f.sets === 'money' ? moneyFormat(rules.currency, true) : f.sets
    return input({ formula: f.formula }, after(f.sets === 'money' ? 'currency' : f.sets === 18 ? 'time' : 'date', sets))
  }
  if (text === '') return cell.numberFormat !== undefined || Object.keys(kept).length ? input({}) : { ok: true, input: null }
  if (text.startsWith("'") && !pasted) return input({ value: text.slice(1) })
  if ((way !== 'replaced' && shows(cell.numberFormat, 'text')) || text === '=') return input({ value: text })
  if (text.startsWith('=')) return formula(text, 0)
  const percent = (way === 'typed' || way === 'inPlace') && shows(cell.numberFormat, 'pct') && (way === 'typed' || /^[-+]/.test(text))
  const typed = readTyped(text, rules, clock, percent)
  // `+` or `-` before what is no constant begins a formula (alone they are text), with the `%` the percent editor added
  // (`-5%` in a percent cell is `-5%%`, the set `entry/typed-keys-held-3`); Excel drops a `+` before a number
  if (/^[-+]./s.test(text) && typed.t === 'text' && typed.format === undefined) return plusNumber(typed.v, rules.decimal) ? formula(typed.v.slice(1), 1) : formula(typed.v, 0)
  switch (typed.t) {
    case 'number':
      // the sign of a locale whose money is not the euro is dropped as typed (`read.ts`)
      if (pasted && typed.form === 'plain' && /^€\d+$/.test(text)) return input({ value: typed.v }, formatAfter('currency', PASTED_EURO, cell.numberFormat, locale))
      return input({ value: typed.v }, after(typed.form, typed.format))
    case 'boolean':
      return input({ value: typed.v })
    case 'text':
      return input({ value: typed.v }, typed.format === undefined ? cell.numberFormat : after('scientific', typed.format))
    case 'error':
      return input({ formula: typed.literal })
  }
}
