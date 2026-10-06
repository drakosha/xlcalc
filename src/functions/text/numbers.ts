/**
 * Numbers as text by the book's locale (`Cx.rules`): FIXED is ported from formula.js 4.6.1 (src/text.js),
 * Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); TEXT and DOLLAR are this engine's own.
 */
import type { Cx, FnDef } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { E, type CErr, type V } from '../../values.js'
import { roundDigits } from '../common.js'
import { roundHalfUp } from '../../format.js'

/** Rounded on the decimal representation (1.005 → 1.01); negative decimals round left of the point; grouped by threes with `group`. */
function fixedDigits(x: number, decimals: number, rules: Rules, group: boolean): { text: string; zero: boolean } {
  const k = Math.max(decimals, 0)
  // toFixed stops at 100 decimals; a double has no further digits
  const text = roundHalfUp(Math.abs(roundDigits(x, decimals, Math.round)), Math.min(k, 100)) + '0'.repeat(Math.max(k - 100, 0))
  const [ip, fp] = text.split('.')
  const grouped = group ? ip.replace(/\B(?=(\d{3})+$)/g, rules.group) : ip
  return { text: grouped + (fp === undefined ? '' : rules.decimal + fp), zero: Number(text) === 0 }
}

export const functions: Record<string, FnDef> = {
  DOLLAR: {
    args: 'n w?',
    cx: true,
    fn: (cx: Cx, x: number, decimals = 2) => {
      if (decimals > 127) return E.VALUE
      const { symbol, before, space, negative } = cx.rules.currency
      const { text } = fixedDigits(x, decimals, cx.rules, true)
      const amount = before ? symbol + space + text : text + space + symbol
      // the en-US currency format writes a negative amount in parentheses, even one rounded to zero
      return x >= 0 ? amount : negative === '()' ? `(${amount})` : `-${amount}`
    },
  },
  FIXED: {
    args: 'n w? b?',
    cx: true,
    fn: (cx: Cx, x: number, decimals = 2, noCommas = false) => {
      if (decimals > 127) return E.VALUE
      const { text, zero } = fixedDigits(x, decimals, cx.rules, !noCommas)
      return (x < 0 && !zero ? '-' : '') + text
    },
  },
  TEXT: { args: 'v s', cx: true, fn: (cx: Cx, v: Exclude<V, CErr>, format: string) => cx.rules.formatText(v ?? 0, format) ?? E.VALUE },
}
