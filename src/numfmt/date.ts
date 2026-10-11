/**
 * Date and time sections: `y`, `m` (minutes after hours or before seconds), `d`, `h`, `s` with fractions of a
 * second (`ss.00`), `AM/PM` and `A/P`, elapsed time (`[h]:mm`, `[mm]:ss`, `[ss]`), names by the locale or a
 * locale tag (`[$-409]`). Excel's calendar from 1900-01-00 to 9999-12-31, or from 1904-01-01 in the 1904 date
 * system (`Style.epoch`), where a negative value shows as its size with a minus; outside it a section that is not
 * elapsed time alone gives null (TEXT answers #VALUE!, a cell shows `#`).
 */
import { ymd } from '../dates.js'
import type { Ph, Section, Tok } from './parse.js'
import { plain, type Names, type Style } from './style.js'

/** The first serial past 9999-12-31. */
const END = 2_958_466

const two = (n: number) => String(n).padStart(2, '0')

/** Whether the `m` at `i` means minutes: after hours or before seconds, across other tokens. */
export function minutes(toks: Tok[], i: number): boolean {
  const unit = (t: Tok | undefined) => (t?.t === 'date' || t?.t === 'elapsed' ? t.s[0] : undefined)
  const prev = unit(toks.slice(0, i).findLast((t) => t.t === 'date' || t.t === 'elapsed'))
  const next = unit(toks.slice(i + 1).find((t) => t.t === 'date' || t.t === 'elapsed'))
  return prev === 'h' || next === 's'
}

/** The number of decimals of a second: `0`s after a point that follows the seconds. */
function subseconds(toks: Tok[], i: number): number {
  let k = 0
  while (toks[i + 1 + k]?.t === 'ph' && (toks[i + 1 + k] as Ph).c === '0') k++
  return k
}

export function formatDate(v: number, sec: Section, style: Style, names: Names): string | null {
  const { toks } = sec
  const epoch = style.epoch ?? 0
  const elapsedOnly = !toks.some((t) => t.t === 'date' || t.t === 'ampm')
  // the 1904 date system shows a negative date or time as its size with a minus (the workbook set date1904)
  if (v < 0 && !elapsedOnly && !epoch) return null
  // the precision of the time: whole seconds, or the decimals after the seconds
  const s = toks.findIndex((t, i) => t.t === 'dot' && (toks[i - 1]?.t === 'date' || toks[i - 1]?.t === 'elapsed') && (toks[i - 1] as { s: string }).s[0] === 's')
  const k = s < 0 ? 0 : Math.min(3, subseconds(toks, s))
  const unit = 86_400 * 10 ** k
  const a = Math.abs(v)
  let day = Math.floor(a)
  let ticks = Math.round((a - day) * unit)
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
    } else if (t.t === 'dot' && i === s && k) {
      out += style.decimal + String(ticks % 10 ** k).padStart(k, '0')
      i += k
    } else if (t.t === 'slash') out += style.dateSeparator
    // separators in a date (but a second's decimal point) are the code's own characters
    else out += t.t === 'dot' || t.t === 'comma' ? t.s : plain(t, style)
  }
  return out
}
