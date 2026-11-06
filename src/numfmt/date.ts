/**
 * Date and time sections: `y`, `m` (minutes after hours or before seconds), `d`, `h`, `s` with fractions of a
 * second (`ss.00`), `AM/PM` and `A/P`, elapsed time (`[h]:mm`, `[mm]:ss`, `[ss]`), names by the locale or a
 * locale tag (`[$-409]`). Excel's calendar from 1900-01-00 to 9999-12-31, or from 1904-01-01 in the 1904 date
 * system (`Style.epoch`), where a negative value shows as its size with a minus; outside it a section that is not
 * elapsed time alone gives null (TEXT answers #VALUE!, a cell shows `#`), and so does a digit placeholder in it.
 */
import { momentOf, ymd } from '../dates.js'
import { positional } from './general.js'
import { hijri } from './hijri.js'
import { dated, misplaced, subseconds, type Section, type Tok } from './parse.js'
import { plain, type Names, type Style } from './style.js'

/** The first serial past 9999-12-31. */
const END = 2_958_466

const two = (n: number) => String(n).padStart(2, '0')

/**
 * Whether the `m` at `i` means minutes: after hours, whatever date's codes stand between (`h dd mm`, the set `links/letters-more`), or before
 * seconds, across tokens other than a date's code but an era's (`mm g ss`), as Microsoft's guide to number formats says, and before a
 * second's decimals (`mm.00`, which the guide leaves to months).
 */
export function minutes(toks: Tok[], i: number): boolean {
  const unit = (t: Tok | undefined) => (dated(t) ? t.s[0] : undefined)
  const prev = unit(toks.slice(0, i).findLast((t) => 'hmns'.includes(unit(t) ?? '-')))
  const subs = subseconds(toks)
  const j = toks.findIndex((t, k) => k > i && ((unit(t) ?? 'g') !== 'g' || subs.has(k)))
  return prev === 'h' || (j > 0 && (subs.has(j) || unit(toks[j]) === 's'))
}

/**
 * What the tokens of a format's section show a number as: a date where a day, a month or a year leads its date tokens,
 * a time of day alone where they are hours, minutes or seconds (with elapsed time and `AM/PM`); undefined for neither.
 */
export function dayOrTime(toks: Tok[]): 'date' | 'time' | undefined {
  const day = (t: Tok, i: number) => t.t === 'date' && !(t.s[0] === 'h' || t.s[0] === 's' || (t.s[0] === 'm' && minutes(toks, i)))
  return toks.some(day) ? 'date' : toks.some((t) => dated(t) || t.t === 'ampm') ? 'time' : undefined
}

/** What a date section's tokens give whatever the value: read once per section, as a column of dates shows one section. */
interface Plan {
  /** Elapsed time alone, which has no calendar's end. */
  elapsedOnly: boolean
  /** The second's decimals (`subseconds`). */
  subs: Map<number, number>
  /** A digit placeholder, a percent or an exponent where a date's code has no place for one, which comes to a section of `AM/PM` alone here. */
  invalid: boolean
  /** The most decimals of a second shown. */
  k: number
  ampm: boolean
  /** By token, whether it is an `m` or `mm` of minutes. */
  minutes: boolean[]
}

/** Plans by a section's tokens, which `parseFormat` keeps one array of per code. */
const PLANS = new WeakMap<Tok[], Plan>()

function planOf(toks: Tok[]): Plan {
  let p = PLANS.get(toks)
  if (p) return p
  const subs = subseconds(toks)
  p = {
    elapsedOnly: !toks.some((t) => t.t === 'date' || t.t === 'ampm'),
    subs,
    invalid: misplaced(toks, subs),
    // the precision of the time: whole seconds, or the most decimals of a second the code shows
    k: Math.max(0, ...[...subs.values()].map((n) => Math.min(3, n))),
    ampm: toks.some((t) => t.t === 'ampm'),
    minutes: toks.map((t, i) => t.t === 'date' && t.s[0] === 'm' && t.s.length <= 2 && minutes(toks, i)),
  }
  PLANS.set(toks, p)
  return p
}

export function formatDate(v: number, sec: Section, style: Style, names: Names): string | null {
  const { toks } = sec
  const epoch = style.epoch ?? 0
  const { elapsedOnly, subs, invalid, k, ampm, minutes: mins } = planOf(toks)
  // the 1904 date system shows a negative date or time as its size with a minus (the workbook set date1904)
  if ((v < 0 && !elapsedOnly && !epoch) || invalid) return null
  const a = Math.abs(v)
  const unit = 86_400 * 10 ** k
  // whole seconds as Excel reads a moment (`momentOf`); decimals of a second round the day's fraction to their unit (TEXT's
  // `ss.000` in the set `seconds-held`, where no case tells this rule from momentOf's at the unit)
  let { day, second: ticks } = k ? { day: Math.floor(a), second: Math.round((a - Math.floor(a)) * unit) } : momentOf(a)
  if (ticks >= unit) [day, ticks] = [day + 1, ticks - unit]
  if (day + epoch >= END && !elapsedOnly) return null
  const secs = Math.floor(ticks / 10 ** k)
  // Excel's calendar: day 0 is 1900-01-00, day 60 the phantom 1900-02-29, day 1 a Sunday; a 1904 book's day 0 is 1462
  const { y: Y, m, d: D } = sec.hijri ? hijri(day + epoch) : ymd(day + epoch)
  const W = (day + epoch + 6) % 7
  const h = Math.floor(secs / 3600)
  const mi = Math.floor(secs / 60) % 60
  const sc = secs % 60
  let out = v < 0 ? '-' : ''
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]
    if (t.t === 'ampm') out += t.s.length > 3 ? (h < 12 ? 'AM' : 'PM') : h < 12 ? t.s[0] : t.s[2]
    else if (t.t === 'elapsed') {
      const total = t.s[0] === 'h' ? day * 24 + h : t.s[0] === 'n' ? (day * 24 + h) * 60 + mi : day * 86_400 + secs
      // digits past 1E+21 too, 15 significant and zeros as a number shows them, not JavaScript's `2.4e+21`
      out += positional(total, 0).padStart(t.s.length, '0')
    } else if (t.t === 'date') {
      const c = t.s[0]
      const n = t.s.length
      if (c === 'g') continue
      if (c === 'y' || c === 'b') {
        const year = c === 'b' ? Y + 543 : Y
        out += n <= 2 ? two(year % 100) : String(year)
      } else if (c === 'd' || c === 'a') out += n === 1 ? String(D) : n === 2 ? two(D) : n === 3 ? names.daysShort[W] : names.days[W]
      else if (c === 'h') {
        const hh = ampm ? h % 12 || 12 : h
        out += n === 1 ? String(hh) : two(hh)
      } else if (c === 's') out += n === 1 ? String(sc) : two(sc)
      else if (c === 'n' || mins[i]) out += n === 1 ? String(mi) : two(mi)
      else out += n === 1 ? String(m) : n === 2 ? two(m) : n === 3 ? names.monthsShort[m - 1] : n === 4 ? names.months[m - 1] : names.months[m - 1][0]
    } else if (subs.has(i)) {
      // each point shows the second's decimals to its own zeros, of the moment rounded to the most of them
      const n = subs.get(i)!
      out += style.decimal + String(ticks % 10 ** k).padStart(k, '0').slice(0, n)
      i += n
    } else if (t.t === 'slash') out += style.dateSeparator
    // separators in a date (but a second's decimal point) are the code's own characters
    else out += t.t === 'dot' || t.t === 'comma' ? t.s : plain(t, style)
  }
  return out
}
