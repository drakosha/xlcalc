/**
 * Text typed into a cell as Excel takes it in a book's locale (`parseEntry`), measured in the entry sets: a formula after
 * `=`, or after a `+` or `-` that begins no number (typed so, Excel makes `+1+2` the formula `1+2` and `-C1` the formula
 * `-C1`), translated from the locale's formula text; text after an apostrophe, or anything in a cell formatted as text;
 * else a constant (`read.ts`), whose form may set the cell's format (`forms.ts`). A number typed into a percent cell
 * gets the `%` Excel's editor adds as it is typed (`5` is 5%). The cell keeps its style and meta. Only a formula that
 * does not parse stops the entry, as Excel's dialog does; an unknown function or name is taken, and the cell computes
 * `#NAME?`.
 */
import type { CellInput, CheckCode, EntryOptions, EntryResult, NumberFormat } from '../contract.js'
import { checkEntryOptions, checkString } from '../input.js'
import { validateLocalFormula, toInvariant } from '../locale/index.js'
import { validate } from '../syntax.js'
import { localeOf } from '../locale/registry.js'
import type { RulesClock } from '../locale/rules/data.js'
import { rulesOf, type Rules } from '../locale/rules.js'
import type { Locale } from '../locale/types.js'
import { epochOf, ymd } from '../dates.js'
import { msToSerial } from '../model.js'
import { formatAfter, shows } from './forms.js'
import { formulaFormat } from './formula.js'
import { moneyFormat, readTyped } from './read.js'

/** The rules of each locale object and date system, with a clock whose year is set for each entry (entries are synchronous). */
const RULES = new WeakMap<Locale, Map<number, { rules: Rules; clock: RulesClock & { y: number } }>>()

function rulesFor(locale: Locale, epoch: number, year: number) {
  let byEpoch = RULES.get(locale)
  if (!byEpoch) RULES.set(locale, (byEpoch = new Map()))
  let r = byEpoch.get(epoch)
  if (!r) {
    const clock: RulesClock & { y: number } = {
      y: year,
      read: false,
      year: () => {
        clock.read = true
        return clock.y
      },
    }
    byEpoch.set(epoch, (r = { rules: rulesOf(locale, clock, epoch), clock }))
  }
  r.clock.y = year
  return r
}

/** The check codes of a formula Excel does not take as typed: its dialog stops the entry. */
const STOPS: ReadonlySet<CheckCode> = new Set(['syntax', 'empty', 'too_long', 'too_deep'])

/**
 * A typed formula, `shift` characters into the typed text: its canonical stored text and the format it sets (`formula.ts`),
 * its invariant text where it names what the book may lack (`#NAME?` then), or why Excel stops it, at its place in the text.
 */
function formulaOf(text: string, shift: number, locale: Locale): { formula: string; sets?: NumberFormat | 'money' } | Extract<EntryResult, { ok: false }> {
  const v = validateLocalFormula(text, locale)
  if (v.ok) {
    const p = validate(v.canonical)
    return { formula: v.canonical, ...(p.ok ? { sets: formulaFormat(p.ast) } : {}) }
  }
  if (STOPS.has(v.code)) return { ok: false, code: v.code, message: v.message, at: v.at + shift }
  const t = toInvariant(text, locale)
  return t.ok ? { formula: t.formula } : { ...t, at: t.at + shift }
}

/** Whether typed text begins with a `+` before a number (`+1*2`, `+.5+1`, `++5`): Excel drops it, the formula is `1*2`. */
const plusNumber = (text: string, decimal: string) => text.startsWith('+') && /^\+?[\d.]/.test(text.slice(1).replace(decimal, '.'))

/**
 * Text typed into a cell as Excel takes it in the book's locale: the cell's new input, with the cell's style and meta
 * and the format Excel leaves it, or a formula Excel does not take. Empty text clears the value and keeps the rest.
 */
export function parseEntry(text: string, opts?: EntryOptions): EntryResult {
  checkString(text, 'text')
  const o = checkEntryOptions(opts)
  const locale = localeOf(o.locale)
  const cell = o.cell ?? {}
  const kept: CellInput = { ...(cell.style !== undefined ? { style: cell.style } : {}), ...(cell.meta !== undefined ? { meta: cell.meta } : {}) }
  const input = (body: CellInput, format: NumberFormat | undefined = cell.format): EntryResult => ({ ok: true, input: { ...body, ...(format !== undefined ? { format } : {}), ...kept } })
  const epoch = epochOf(o.dateSystem)
  const year = ymd(Math.floor(msToSerial(o.now === undefined ? Date.now() : Date.parse(o.now)))).y
  const { rules, clock } = rulesFor(locale, epoch, year)
  // a formula takes the format of what it computes as a typed constant of that kind would (`forms.ts`)
  const formula = (src: string, shift: number): EntryResult => {
    const f = formulaOf(src, shift, locale)
    if (!('formula' in f)) return f
    if (f.sets === undefined) return input({ formula: f.formula })
    const sets = f.sets === 'money' ? moneyFormat(rules.currency, true) : f.sets
    return input({ formula: f.formula }, formatAfter(f.sets === 'money' ? 'currency' : f.sets === 18 ? 'time' : 'date', sets, cell.format, locale))
  }
  if (text === '') return cell.format !== undefined || Object.keys(kept).length ? input({}) : { ok: true, input: null }
  if (text.startsWith("'")) return input({ value: text.slice(1) })
  if (shows(cell.format, 'text') || text === '=') return input({ value: text })
  if (text.startsWith('=')) return formula(text, 0)
  const typed = readTyped(text, rules, clock, shows(cell.format, 'pct'))
  // `+` or `-` before what is no constant begins a formula (alone they are text); Excel drops a `+` before a number
  if (/^[-+]./s.test(text) && typed.t === 'text' && typed.format === undefined) return plusNumber(text, rules.decimal) ? formula(text.slice(1), 1) : formula(text, 0)
  switch (typed.t) {
    case 'number':
      return input({ value: typed.v }, formatAfter(typed.form, typed.format, cell.format, locale))
    case 'boolean':
      return input({ value: typed.v })
    case 'text':
      return input({ value: typed.v }, typed.format === undefined ? cell.format : formatAfter('scientific', typed.format, cell.format, locale))
    case 'error':
      return input({ formula: typed.literal })
  }
}
