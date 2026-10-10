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
import { DEFAULT_LOCALE, LOCALE_FILES } from '../locale/registry.js'
import type { RulesClock } from '../locale/rules/data.js'
import { rulesOf, type Rules } from '../locale/rules.js'
import { epochOf, ymd } from '../dates.js'
import { msToSerial } from '../model.js'
import { literal } from '../criteria.js'
import { builtinOf, formatAfter, shows } from './forms.js'
import { formulaFormat } from './formula.js'
import { moneyFormat, readTyped } from './read.js'

/** The rules of each locale and date system, with a clock whose year is set for each entry (entries are synchronous). */
const RULES = new Map<string, { rules: Rules; clock: RulesClock & { y: number } }>()

function rulesFor(tag: string, epoch: number, year: number) {
  const k = `${tag}|${epoch}`
  let r = RULES.get(k)
  if (!r) {
    const clock: RulesClock & { y: number } = {
      y: year,
      read: false,
      year: () => {
        clock.read = true
        return clock.y
      },
    }
    RULES.set(k, (r = { rules: rulesOf(tag, clock, epoch), clock }))
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
function formulaOf(text: string, shift: number, tag: string): { formula: string; sets?: NumberFormat | 'money' } | Extract<EntryResult, { ok: false }> {
  const locale = LOCALE_FILES.get(tag)!.text
  const v = validateLocalFormula(text, locale)
  if (v.ok) {
    const p = validate(v.canonical)
    return { formula: v.canonical, ...(p.ok ? { sets: formulaFormat(p.ast) } : {}) }
  }
  if (STOPS.has(v.code)) return { ok: false, code: v.code, message: v.message, at: v.at + shift }
  const t = toInvariant(text, locale)
  return t.ok ? { formula: t.formula } : { ...t, at: t.at + shift }
}

/**
 * Text typed into a cell as Excel takes it in the book's locale: the cell's new input, with the cell's style and meta
 * and the format Excel leaves it, or a formula Excel does not take. Empty text clears the value and keeps the rest.
 */
export function parseEntry(text: string, opts?: EntryOptions): EntryResult {
  checkString(text, 'text')
  const o = checkEntryOptions(opts)
  const tag = o.locale ?? DEFAULT_LOCALE
  const cell = o.cell ?? {}
  const kept: CellInput = { ...(cell.style !== undefined ? { style: cell.style } : {}), ...(cell.meta !== undefined ? { meta: cell.meta } : {}) }
  const input = (body: CellInput, format: NumberFormat | undefined = cell.format): EntryResult => ({ ok: true, input: { ...body, ...(format !== undefined ? { format } : {}), ...kept } })
  const epoch = epochOf(o.dateSystem)
  const year = ymd(Math.floor(msToSerial(o.now === undefined ? Date.now() : Date.parse(o.now)))).y
  const { rules, clock } = rulesFor(tag, epoch, year)
  // a formula in a cell in General takes the format of what it computes
  const formula = (src: string, shift: number): EntryResult => {
    const f = formulaOf(src, shift, tag)
    if (!('formula' in f)) return f
    const general = cell.format === undefined || builtinOf(cell.format, tag) === 0
    return input({ formula: f.formula }, general && f.sets !== undefined ? (f.sets === 'money' ? moneyFormat(rules.currency, true) : f.sets) : cell.format)
  }
  if (text === '') return cell.format !== undefined || Object.keys(kept).length ? input({}) : { ok: true, input: null }
  if (text.startsWith("'")) return input({ value: text.slice(1) })
  if (shows(cell.format, 'text') || text === '=') return input({ value: text })
  if (text.startsWith('=')) return formula(text, 0)
  // the editor of a percent cell adds `%` after a number as it is typed
  const percent = shows(cell.format, 'pct') && new RegExp(`^[-+]?(\\d|${literal(rules.decimal)})`).test(text)
  const typed = readTyped(percent ? `${text}%` : text, rules, clock)
  // `+` or `-` before what is no number begins a formula; `+` before a digit gives its place to `=`
  if ((text.startsWith('+') || text.startsWith('-')) && typed.t !== 'number') return /^\+\d/.test(text) ? formula(text.slice(1), 1) : formula(text, 0)
  switch (typed.t) {
    case 'number':
      return input({ value: typed.v }, formatAfter(typed.form, typed.format, cell.format, tag))
    case 'boolean':
      return input({ value: typed.v })
    case 'text':
      return input({ value: typed.v }, typed.format === undefined ? cell.format : formatAfter('scientific', typed.format, cell.format, tag))
    case 'error':
      return input({ formula: typed.literal })
  }
}
