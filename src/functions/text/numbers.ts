/**
 * Numbers as text by the book's locale (`Cx.rules`): FIXED is ported from formula.js 4.6.1 (src/text.js),
 * Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); TEXT and DOLLAR are this engine's own.
 */
import { positional } from '../../numfmt/general.js'
import { MAX_SHOWN } from '../../numfmt/glyphs.js'
import type { Cx, FnDef } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { E, type CErr, type V } from '../../values.js'
import { roundDigits } from '../common.js'

/**
 * Rounded on the decimal representation (1.005 → 1.01); negative decimals round left of the point; 15 significant
 * digits and zeros past them, as TEXT writes them (`positional`); grouped by threes with `group`.
 */
function fixedDigits(x: number, decimals: number, rules: Rules, group: boolean): { text: string; zero: boolean } {
  const k = Math.max(decimals, 0)
  const text = positional(Math.abs(roundDigits(x, decimals, Math.round)), k)
  const [ip, fp] = text.split('.')
  const grouped = group ? ip.replace(/\B(?=(\d{3})+$)/g, rules.group) : ip
  return { text: grouped + (fp === undefined ? '' : rules.decimal + fp), zero: Number(text) === 0 }
}

/** Text past `room` characters is #VALUE!, as TEXT's past 255 (`FIXED(1E+200,0)`, the set `values-held`). */
const shown = (text: string, room = MAX_SHOWN): V => (text.length > room ? E.VALUE : text)

export const functions: Record<string, FnDef> = {
  DOLLAR: {
    args: 'n w?',
    cx: true,
    fn: (cx: Cx, x: number, decimals = 2) => {
      if (decimals > 127) return E.VALUE
      const { symbol, before, space, negative } = cx.rules.currency
      const { text } = fixedDigits(x, decimals, cx.rules, true)
      const amount = before ? symbol + space + text : text + space + symbol
      // the en-US currency format writes a negative amount in parentheses, even one rounded to zero, and pads a
      // positive one with the space of `_)`, which counts toward 255 characters (`DOLLAR(1E+190,0)`, the set `values-held-2`)
      if (x >= 0) return shown(amount, negative === '()' ? MAX_SHOWN - 1 : MAX_SHOWN)
      return shown(negative === '()' ? `(${amount})` : `-${amount}`)
    },
  },
  FIXED: {
    args: 'n w? b?',
    cx: true,
    fn: (cx: Cx, x: number, decimals = 2, noCommas = false) => {
      if (decimals > 127) return E.VALUE
      const { text, zero } = fixedDigits(x, decimals, cx.rules, !noCommas)
      return shown((x < 0 && !zero ? '-' : '') + text)
    },
  },
  TEXT: { args: 'v s', cx: true, fn: (cx: Cx, v: Exclude<V, CErr>, format: string) => cx.rules.formatText(v ?? 0, format) ?? E.VALUE },
}
