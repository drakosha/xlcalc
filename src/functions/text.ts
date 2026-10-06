/**
 * Text functions.
 *
 * CLEAN … UNICODE are ported from formula.js 4.6.1 (https://github.com/formulajs/formulajs), src/text.js:
 * Copyright (c) 2014 Sutoiku, Inc., MIT License (full text in THIRD_PARTY.md). Changed: arguments are
 * coerced by the engine's argument layer (`src/args.ts`; numbers become text in the General format);
 * results longer than Excel's 32,767 characters are #VALUE!; FIXED rounds on the decimal representation
 * (1.005 → 1.01); LEFT and RIGHT with a negative count are #VALUE!; FIND checks the start position; TRIM
 * removes only spaces (Excel keeps tabs, line breaks and non-breaking spaces; formula.js collapsed all
 * whitespace); UNICODE reads a whole code point (formula.js took the first UTF-16 unit).
 * TEXT is this engine's own (`src/format.ts`).
 */
import { eachValue, type FnDef, type Segment } from '../args.js'
import { formatText, roundHalfUp } from '../format.js'
import { E, isErr, toText, type CErr, type V } from '../values.js'
import { roundDigits } from './common.js'

/** Excel's limit on the length of a text value. */
const MAX_TEXT = 32_767

const capped = (s: string): V => (s.length > MAX_TEXT ? E.VALUE : s)

export const functions: Record<string, FnDef> = {
  CLEAN: { args: 's', fn: (s: string) => s.replace(/[\0-\x1F]/g, '') },
  // ranges in order, a blank cell as ""
  CONCAT: {
    args: 'V',
    rep: 1,
    fn: (segs: Segment[]) => {
      let out = ''
      const e = eachValue(segs, (v) => {
        const t = toText(v)
        if (isErr(t)) return t
        out += t
      })
      return e ?? capped(out)
    },
  },
  CONCATENATE: { args: 's', rep: 1, fn: (texts: string[]) => capped(texts.join('')) },
  EXACT: { args: 's s', fn: (a: string, b: string) => a === b },
  FIND: {
    args: 's s i?',
    fn: (find: string, within: string, start = 1) => {
      if (start < 1 || start > within.length) return E.VALUE
      const at = within.indexOf(find, start - 1)
      return at < 0 ? E.VALUE : at + 1
    },
  },
  FIXED: {
    args: 'n i? b?',
    fn: (x: number, decimals = 2, noCommas = false) => {
      if (decimals > 127) return E.VALUE
      const k = Math.max(decimals, 0)
      const text = roundHalfUp(Math.abs(roundDigits(x, decimals, Math.round)), k)
      const [ip, fp] = text.split('.')
      const grouped = noCommas ? ip : ip.replace(/\B(?=(\d{3})+$)/g, ',')
      const negative = x < 0 && Number(text) !== 0
      return (negative ? '-' : '') + grouped + (fp === undefined ? '' : '.' + fp)
    },
  },
  LEFT: { args: 's i?', fn: (s: string, n = 1) => (n < 0 ? E.VALUE : s.substring(0, n)) },
  LEN: { args: 's', fn: (s: string) => s.length },
  LOWER: { args: 's', fn: (s: string) => s.toLowerCase() },
  REPT: { args: 's i', fn: (s: string, n: number) => (n < 0 || s.length * n > MAX_TEXT ? E.VALUE : s.repeat(n)) },
  RIGHT: { args: 's i?', fn: (s: string, n = 1) => (n < 0 ? E.VALUE : s.substring(Math.max(0, s.length - n))) },
  T: { args: 'a', fn: (v: V) => (typeof v === 'string' || isErr(v) ? v : '') },
  TEXT: {
    args: 'v s',
    fn: (v: Exclude<V, CErr>, format: string) => formatText(v ?? 0, format) ?? E.VALUE,
  },
  TRIM: { args: 's', fn: (s: string) => s.replace(/ +/g, ' ').replace(/^ | $/g, '') },
  UNICODE: { args: 's', fn: (s: string) => (s === '' ? E.VALUE : s.codePointAt(0)!) },
}
