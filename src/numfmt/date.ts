/**
 * Date and time sections: `y`, `m` (minutes after hours or before seconds), `d`, `h`, `s` with fractions of a
 * second (`ss.00`), `AM/PM` and `A/P`, elapsed time (`[h]:mm`, `[mm]:ss`, `[ss]`), names by the locale or a
 * locale tag (`[$-409]`). Excel's calendar from 1900-01-00 to 9999-12-31, or from 1904-01-01 in the 1904 date
 * system (`Style.epoch`), where a negative value shows as its size with a minus; outside it a section that is not
 * elapsed time alone gives null (TEXT answers #VALUE!, a cell shows `#`), and so does a digit placeholder in it.
 */
import { momentOf, ymd } from '../dates.js'
import type { Ph, Section, Tok } from './parse.js'
import { plain, type Names, type Style } from './style.js'

/** The first serial past 9999-12-31. */
const END = 2_958_466

const two = (n: number) => String(n).padStart(2, '0')

/**
 * The second's decimals of a date's code as [MS-XLS] 2.4.126 gives them (`NFPartSubSecond`, the grammar of ECMA-376
 * 18.8.31's codes): a point and one to three zeros anywhere after the first date or time code (`ss.00`, `dd .000`,
 * `yy.00.00`, `h AM/PM.0`), by the point's position and its zeros.
 */
function subseconds(toks: Tok[]): Map<number, number> {
  const out = new Map<number, number>()
  let dated = false
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]
    if (t.t === 'date' || t.t === 'elapsed') dated = true
    else if ((t.t === 'dot' || (t.t === 'lit' && t.point)) && dated && toks[i + 1]?.t === 'ph' && (toks[i + 1] as Ph).c === '0') {
      let k = 0
      while (toks[i + 1 + k]?.t === 'ph' && (toks[i + 1 + k] as Ph).c === '0') k++
      out.set(i, k)
      i += k
    }
  }
  return out
}

/**
 * Whether the `m` at `i` means minutes: after hours or before seconds, across other tokens, as Microsoft's guide
 * to number formats says, and before a second's decimals (`mm.00`, which the guide leaves to months).
 */
export function minutes(toks: Tok[], i: number): boolean {
  const unit = (t: Tok | undefined) => (t?.t === 'date' || t?.t === 'elapsed' ? t.s[0] : undefined)
  const prev = unit(toks.slice(0, i).findLast((t) => t.t === 'date' || t.t === 'elapsed'))
  const subs = subseconds(toks)
  const j = toks.findIndex((t, k) => k > i && (t.t === 'date' || t.t === 'elapsed' || subs.has(k)))
  return prev === 'h' || (j > 0 && (subs.has(j) || unit(toks[j]) === 's'))
}

export function formatDate(v: number, sec: Section, style: Style, names: Names): string | null {
  const { toks } = sec
  const epoch = style.epoch ?? 0
  const elapsedOnly = !toks.some((t) => t.t === 'date' || t.t === 'ampm')
  // the 1904 date system shows a negative date or time as its size with a minus (the workbook set date1904)
  if (v < 0 && !elapsedOnly && !epoch) return null
  // the precision of the time: whole seconds, or the most decimals of a second the code shows
  const subs = subseconds(toks)
  // a digit placeholder other than the second's decimals, a percent or an exponent has no place in a date's code
  // (`NFDateTime`), which is then invalid: TEXT(1,"mm###") is #VALUE!, and so is a fourth zero of decimals
  const decimal = (i: number) => [...subs].some(([at, n]) => i > at && i <= at + Math.min(3, n))
  if (toks.some((t, i) => (t.t === 'ph' && !decimal(i)) || t.t === 'pct' || t.t === 'exp')) return null
  const k = Math.max(0, ...[...subs.values()].map((n) => Math.min(3, n)))
  const a = Math.abs(v)
  const unit = 86_400 * 10 ** k
  // whole seconds as Excel reads a moment (`momentOf`); decimals of a second round the day's fraction to their unit (TEXT's
  // `ss.000` in the set `seconds-held`, where no case tells this rule from momentOf's at the unit)
  let { day, second: ticks } = k ? { day: Math.floor(a), second: Math.round((a - Math.floor(a)) * unit) } : momentOf(a)
  if (ticks >= unit) [day, ticks] = [day + 1, ticks - unit]
  if (day + epoch >= END && !elapsedOnly) return null
  const secs = Math.floor(ticks / 10 ** k)
  const ampm = toks.some((t) => t.t === 'ampm')
  // Excel's calendar: day 0 is 1900-01-00, day 60 the phantom 1900-02-29, day 1 a Sunday; a 1904 book's day 0 is 1462
  const { y: Y, m, d: D } = ymd(day + epoch)
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
      out += String(total).padStart(t.s.length, '0')
    } else if (t.t === 'date') {
      const c = t.s[0]
      const n = t.s.length
      if (c === 'y') out += n <= 2 ? two(Y % 100) : String(Y)
      else if (c === 'd') out += n === 1 ? String(D) : n === 2 ? two(D) : n === 3 ? names.daysShort[W] : names.days[W]
      else if (c === 'h') {
        const hh = ampm ? h % 12 || 12 : h
        out += n === 1 ? String(hh) : two(hh)
      } else if (c === 's') out += n === 1 ? String(sc) : two(sc)
      else if (c === 'n' || (n <= 2 && minutes(toks, i))) out += n === 1 ? String(mi) : two(mi)
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
