/** VALUE and NUMBERVALUE, this engine's own: text read as Excel en-US reads it, checked case by case against Excel. */
import type { FnDef } from '../../args.js'
import { parseDateText } from '../../dates.js'
import { readNumber } from '../../model.js'
import { E, type CErr, type V } from '../../values.js'

/** Digits with commas, each followed by three digits at least (Excel: `1,0000` is 10000, `10,00` and `0,123` are not numbers). */
const GROUPED = /^((?!0+,)\d+(?:,\d{3,})*|)(\.\d*)?$/

/** Excel trims spaces only, not tabs or line breaks. */
const trimSpaces = (s: string) => s.replace(/^ +| +$/g, '')

/** VALUE's time: minutes and seconds past 59 carry (`1:61`), up to 9,999 hours; with ` AM`/` PM` a clock time. */
const TIME = /^(\d+):(\d+)(?::(\d+(?:\.\d+)?))?(?: ([ap]m))?$/i
const MAX_TIME = 10_000 / 24

function timeOf(s: string): number | null {
  const m = TIME.exec(s)
  if (!m) return null
  let h = Number(m[1])
  const mi = Number(m[2])
  const sec = Number(m[3] ?? 0)
  if (m[4]) {
    if (h > 12 || mi > 59 || sec >= 60) return null
    h = (h % 12) + (m[4].toLowerCase() === 'pm' ? 12 : 0)
  }
  const t = (h * 3600 + mi * 60 + sec) / 86_400
  return t < MAX_TIME ? t : null
}

/** Date text, optionally followed by a time that may run past 24 hours (`2025-01-31 25:00`). */
function dateTime(s: string): number | null {
  const m = /^(.*\S) +(\d+:\d+(?::\d+(?:\.\d+)?)?(?: [ap]m)?)$/i.exec(s)
  if (m) {
    const d = m[1].includes(':') ? null : parseDateText(m[1])
    const t = timeOf(m[2])
    return d === null || t === null ? null : d + t
  }
  return s.includes(':') || /\d[ap]m$/i.test(s) ? null : parseDateText(s)
}

/** Number text: sign, `$`, parentheses, commas, exponent, one `%` at either end, `1 1/2`; else a time or a date. */
export function valueOf(text: string): number | null {
  let s = trimSpaces(text)
  if (s === '') return null
  let paren = false
  let percent = false
  // parentheses and the percent sign in either order: `(5%)`, `(5)%`, `%(5)`
  for (;;) {
    if (!paren && s.startsWith('(') && s.endsWith(')')) [paren, s] = [true, s.slice(1, -1)]
    else if (!percent && s.startsWith('%')) [percent, s] = [true, s.slice(1)]
    else if (!percent && s.endsWith('%')) [percent, s] = [true, s.slice(0, -1)]
    else break
    s = trimSpaces(s)
  }
  const sign = /^([+-]?) *(\$?) *([+-]?) *(.*)$/.exec(s)!
  if (sign[1] && sign[3]) return null
  const negative = (sign[1] || sign[3]) === '-'
  if (paren && negative) return null
  const body = sign[4]
  const exp = /^(.*?)([eE][+-]?\d+)$/.exec(body)
  const mant = exp ? exp[1] : body
  let n: number | null = null
  if (GROUPED.test(mant) && /\d/.test(mant)) n = readNumber(mant.replace(/,/g, '') + (exp ? exp[2] : ''))
  else if (!sign[2] && !paren) {
    const frac = /^(\d+) (\d+)\/(\d+)$/.exec(body)
    if (frac && Number(frac[3]) > 0) n = Number(frac[1]) + Number(frac[2]) / Number(frac[3])
    else if (!sign[1] && !sign[3] && !percent) return timeOf(s) ?? dateTime(s)
  }
  if (n === null || !Number.isFinite(n)) return null
  if (percent) n /= 100
  return negative || paren ? -n : n
}

function value(v: Exclude<V, CErr>): V {
  if (typeof v === 'number') return v
  if (v === null) return 0
  if (typeof v === 'boolean') return E.VALUE
  return valueOf(v) ?? E.VALUE
}

/**
 * Group separators go first (after the decimal one they are #VALUE!), the decimal separator becomes a point
 * and the rest reads as in VALUE; whitespace anywhere is dropped and each trailing `%` divides by 100.
 */
function numbervalue(text: string, decimal?: string, group?: string): V {
  if (decimal === '' || group === '') return E.VALUE
  const d = decimal?.[0] ?? '.'
  // a comma as the decimal separator makes the point the default group separator
  const g = group?.[0] ?? (d === ',' ? '.' : ',')
  if (d === g) return E.VALUE
  let s = text.replace(/[ \t\n\r]/g, '')
  if (s === '') return 0
  let pct = 0
  for (; s.endsWith('%'); pct++) s = s.slice(0, -1)
  const at = s.indexOf(d)
  const frac = at < 0 ? '' : s.slice(at + 1)
  if (frac.includes(d) || frac.includes(g)) return E.VALUE
  const n = valueOf((at < 0 ? s : s.slice(0, at)).split(g).join('') + (at < 0 ? '' : '.' + frac))
  return n === null ? E.VALUE : n / 100 ** pct
}

export const functions: Record<string, FnDef> = {
  NUMBERVALUE: { args: 's s? s?', fn: numbervalue },
  VALUE: { args: 'v', fn: value },
}
