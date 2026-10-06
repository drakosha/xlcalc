/**
 * Numbers as text: FIXED is ported from formula.js 4.6.1 (src/text.js), Copyright (c) 2014 Sutoiku, Inc.,
 * MIT License (THIRD_PARTY.md); TEXT and DOLLAR are this engine's own.
 */
import { formatText, roundHalfUp } from '../../format.js'
import type { FnDef } from '../../args.js'
import { E, type CErr, type V } from '../../values.js'
import { roundDigits } from '../common.js'

/** Rounded on the decimal representation (1.005 → 1.01); negative decimals round left of the point. */
function fixedDigits(x: number, decimals: number, commas: boolean): { text: string; zero: boolean } {
  const k = Math.max(decimals, 0)
  // toFixed stops at 100 decimals; a double has no further digits
  const text = roundHalfUp(Math.abs(roundDigits(x, decimals, Math.round)), Math.min(k, 100)) + '0'.repeat(Math.max(k - 100, 0))
  const [ip, fp] = text.split('.')
  const grouped = commas ? ip.replace(/\B(?=(\d{3})+$)/g, ',') : ip
  return { text: grouped + (fp === undefined ? '' : '.' + fp), zero: Number(text) === 0 }
}

export const functions: Record<string, FnDef> = {
  DOLLAR: {
    args: 'n w?',
    fn: (x: number, decimals = 2) => {
      if (decimals > 127) return E.VALUE
      const { text } = fixedDigits(x, decimals, true)
      // the en-US currency format writes a negative amount in parentheses, even one rounded to zero
      return x < 0 ? `($${text})` : `$${text}`
    },
  },
  FIXED: {
    args: 'n w? b?',
    fn: (x: number, decimals = 2, noCommas = false) => {
      if (decimals > 127) return E.VALUE
      const { text, zero } = fixedDigits(x, decimals, !noCommas)
      return (x < 0 && !zero ? '-' : '') + text
    },
  },
  TEXT: { args: 'v s', fn: (v: Exclude<V, CErr>, format: string) => formatText(v ?? 0, format) ?? E.VALUE },
}
