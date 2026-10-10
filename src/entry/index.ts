/**
 * Text typed into a cell as Excel takes it in a book's locale (`parseEntry`): a formula after `=` (or a leading `+`,
 * or `-` before what is no number), translated from the locale's formula text; text after an apostrophe, or anything in
 * a cell formatted as text; else a constant (`read.ts`), whose form may set the cell's format (`forms.ts`). The cell
 * keeps its style and meta. Only a formula that does not parse stops the entry, as Excel's dialog does; an unknown
 * function or name is taken, and the cell computes `#NAME?`.
 */
import type { CellInput, CheckCode, EntryOptions, EntryResult, NumberFormat } from '../contract.js'
import { checkEntryOptions, checkString } from '../input.js'
import { validateLocalFormula, toInvariant } from '../locale/index.js'
import { DEFAULT_LOCALE, LOCALE_FILES } from '../locale/registry.js'
import type { RulesClock } from '../locale/rules/data.js'
import { rulesOf, type Rules } from '../locale/rules.js'
import { epochOf, ymd } from '../dates.js'
import { msToSerial } from '../model.js'
import { category, formatAfter, percentEntry } from './forms.js'
import { readTyped } from './read.js'

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

/** A typed formula: its canonical stored text, its invariant text where it names what the book may lack (`#NAME?` then), or why Excel stops it. */
function formulaOf(text: string, tag: string): { formula: string } | Extract<EntryResult, { ok: false }> {
  const locale = LOCALE_FILES.get(tag)!.text
  const v = validateLocalFormula(text, locale)
  if (v.ok) return { formula: v.canonical }
  if (STOPS.has(v.code)) return { ok: false, code: v.code, message: v.message, at: v.at }
  const t = toInvariant(text, locale)
  return t.ok ? { formula: t.formula } : t
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
  if (text === '') return cell.format !== undefined || Object.keys(kept).length ? input({}) : { ok: true, input: null }
  if (text.startsWith("'")) return input({ value: text.slice(1) })
  if (category(cell.format) === 'text') return input({ value: text })
  const epoch = epochOf(o.dateSystem)
  const year = ymd(Math.floor(msToSerial(o.now === undefined ? Date.now() : Date.parse(o.now)))).y
  const { rules, clock } = rulesFor(tag, epoch, year)
  const typed = readTyped(text, rules, clock)
  // `=` begins a formula, so does `+`, and `-` before what is no number
  if (text.startsWith('=') || text.startsWith('+') || (text.startsWith('-') && typed.t !== 'number')) {
    const f = formulaOf(text, tag)
    return 'formula' in f ? input(f) : f
  }
  switch (typed.t) {
    case 'number':
      return input({ value: percentEntry(typed, cell.format) }, formatAfter(typed, cell.format))
    case 'boolean':
    case 'text':
      return input({ value: typed.v })
    case 'error':
      return input({ formula: typed.literal })
  }
}
