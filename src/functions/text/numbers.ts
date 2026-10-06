/**
 * Numbers and text: FIXED is ported from formula.js 4.6.1 (src/text.js), Copyright (c) 2014 Sutoiku, Inc.,
 * MIT License (THIRD_PARTY.md); TEXT, DOLLAR, VALUE and NUMBERVALUE are this engine's own.
 */
import type { FnDef } from '../../args.js'
import { parseDateText } from '../../dates.js'
import { formatText, roundHalfUp } from '../../format.js'
import { NUMBER_TEXT, readNumber } from '../../model.js'
import { E, type CErr, type V } from '../../values.js'
import { roundDigits } from '../common.js'

/** Rounded on the decimal representation (1.005 → 1.01); negative decimals round left of the point. */
function fixedDigits(x: number, decimals: number, commas: boolean): { text: string; negative: boolean } {
  const k = Math.max(decimals, 0)
  const text = roundHalfUp(Math.abs(roundDigits(x, decimals, Math.round)), k)
  const [ip, fp] = text.split('.')
  const grouped = commas ? ip.replace(/\B(?=(\d{3})+$)/g, ',') : ip
  return { text: grouped + (fp === undefined ? '' : '.' + fp), negative: x < 0 && Number(text) !== 0 }
}

/** Digits with thousands groups of three (`1,234.5`; Excel en-US rejects `1,5`). */
const GROUPED = /^(\d{1,3}(?:,\d{3})+|\d*)(\.\d*)?$/

/** Number text as Excel en-US reads it in VALUE: sign, `$`, parentheses, groups, exponent, `%`, `1 1/2`, else date or time text. */
export function valueOf(text: string): number | null {
  let s = text.trim()
  if (s === '') return null
  let paren = false
  if (s.startsWith('(') && s.endsWith(')')) {
    paren = true
    s = s.slice(1, -1).trim()
  }
  let percent = false
  if (s.endsWith('%')) {
    percent = true
    s = s.slice(0, -1).trim()
  }
  const sign = /^([+-]?)\s*(\$?)\s*([+-]?)\s*(.*)$/.exec(s)!
  if (sign[1] && sign[3]) return null
  const negative = (sign[1] || sign[3]) === '-'
  if (paren && (negative || percent)) return null
  const body = sign[4]
  const exp = /^(.*?)([eE][+-]?\d+)$/.exec(body)
  const mant = exp ? exp[1] : body
  let n: number | null = null
  if (GROUPED.test(mant) && /\d/.test(mant)) n = readNumber(mant.replace(/,/g, '') + (exp ? exp[2] : ''))
  else if (!sign[2] && !paren) {
    const frac = /^(\d+) (\d+)\/(\d+)$/.exec(body)
    if (frac && Number(frac[3]) > 0) n = Number(frac[1]) + Number(frac[2]) / Number(frac[3])
    else if (!sign[1] && !sign[3] && !percent) return parseDateText(text)
  }
  if (n === null) return null
  if (percent) n /= 100
  return negative || paren ? -n : n
}

function value(v: Exclude<V, CErr>): V {
  if (typeof v === 'number') return v
  if (v === null) return 0
  if (typeof v === 'boolean') return E.VALUE
  return valueOf(v) ?? E.VALUE
}

/** Separators are the first characters of their arguments; spaces anywhere are dropped, each trailing `%` divides by 100. */
function numbervalue(text: string, decimal = '.', group = ','): V {
  if (decimal === '' || group === '') return E.VALUE
  const d = decimal[0]
  const g = group[0]
  if (d === g) return E.VALUE
  let s = text.replace(/\s+/g, '')
  if (s === '') return 0
  let pct = 0
  for (; s.endsWith('%'); pct++) s = s.slice(0, -1)
  const at = s.indexOf(d)
  const frac = at < 0 ? '' : s.slice(at + 1)
  if (frac.includes(d) || frac.includes(g)) return E.VALUE
  const t = (at < 0 ? s : s.slice(0, at)).split(g).join('') + (at < 0 ? '' : '.' + frac)
  return NUMBER_TEXT.test(t) ? readNumber(t) / 100 ** pct : E.VALUE
}

export const functions: Record<string, FnDef> = {
  DOLLAR: {
    args: 'n i?',
    fn: (x: number, decimals = 2) => {
      if (decimals > 127) return E.VALUE
      const { text, negative } = fixedDigits(x, decimals, true)
      // the en-US currency format writes a negative amount in parentheses
      return negative ? `($${text})` : `$${text}`
    },
  },
  FIXED: {
    args: 'n i? b?',
    fn: (x: number, decimals = 2, noCommas = false) => {
      if (decimals > 127) return E.VALUE
      const { text, negative } = fixedDigits(x, decimals, !noCommas)
      return (negative ? '-' : '') + text
    },
  },
  NUMBERVALUE: { args: 's s? s?', fn: numbervalue },
  TEXT: { args: 'v s', fn: (v: Exclude<V, CErr>, format: string) => formatText(v ?? 0, format) ?? E.VALUE },
  VALUE: { args: 'v', fn: value },
}
