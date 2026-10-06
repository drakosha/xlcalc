/**
 * Text functions on strings. Ported from formula.js 4.6.1 (src/text.js), Copyright (c) 2014 Sutoiku, Inc.,
 * MIT License (THIRD_PARTY.md); fixed against Excel en-US, see the comments.
 */
import { eachValue, type FnDef, type Segment } from '../../args.js'
import { wildSource } from '../../criteria.js'
import { E, isErr, toText, type V } from '../../values.js'
import { capped, MAX_TEXT, texts } from './common.js'

/** The largest start position MID and REPLACE accept. */
const MAX_POS = 2_147_483_647

/** Case mapping that keeps a character whose mapping is longer (Excel: UPPER("ß") is "ß"). */
function caseOf(ch: string, upper: boolean): string {
  const c = upper ? ch.toUpperCase() : ch.toLowerCase()
  return c.length === ch.length ? c : ch
}

function upper(s: string): string {
  const u = s.toUpperCase()
  if (u.length === s.length) return u
  let out = ''
  for (const ch of s) out += caseOf(ch, true)
  return out
}

/** Excel capitalizes every letter that follows a non-letter, in any alphabet ("2-way" → "2-Way"). */
function proper(s: string): string {
  let out = ''
  let inWord = false
  for (const ch of s) {
    const letter = /\p{L}/u.test(ch)
    out += letter ? caseOf(ch, !inWord) : ch
    inWord = letter
  }
  return out
}

/** Wildcards and case as in the criteria (`src/criteria.ts`). */
function search(find: string, within: string, start = 1): V {
  if (start < 1 || start > within.length) return E.VALUE
  if (find === '') return start
  const m = new RegExp(wildSource(find), 'iu').exec(within.slice(start - 1))
  return m ? m.index + start : E.VALUE
}

function substitute(text: string, old: string, rep: string, instance?: number): V {
  if (instance !== undefined && instance < 1) return E.VALUE
  if (old === '') return text
  if (instance === undefined) return capped(text.split(old).join(rep))
  let at = -old.length
  for (let k = 0; k < instance; k++) {
    at = text.indexOf(old, at + old.length)
    if (at < 0) return text
  }
  return capped(text.slice(0, at) + rep + text.slice(at + old.length))
}

/** Several delimiters are used in turn; ranges are read row by row. */
function textjoin(delims: Segment, ignoreEmpty: boolean, segs: Segment[]): V {
  const ds = texts(delims)
  if (isErr(ds)) return ds
  const parts: string[] = []
  const e = eachValue(segs, (v) => {
    const t = toText(v)
    if (isErr(t)) return t
    if (!(ignoreEmpty && t === '')) parts.push(t)
  })
  if (e) return e
  let out = parts[0] ?? ''
  for (let i = 1; i < parts.length; i++) out += ds[(i - 1) % ds.length] + parts[i]
  return capped(out)
}

export const functions: Record<string, FnDef> = {
  // Excel removes only codes 0-31
  CLEAN: { args: 's', fn: (s: string) => s.replace(/[\0-\x1F]/g, '') },
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
  CONCATENATE: { args: 's', rep: 1, fn: (ts: string[]) => capped(ts.join('')) },
  EXACT: { args: 's s', fn: (a: string, b: string) => a === b },
  FIND: {
    args: 's s i?',
    fn: (find: string, within: string, start = 1) => {
      if (start < 1 || start > within.length) return E.VALUE
      const at = within.indexOf(find, start - 1)
      return at < 0 ? E.VALUE : at + 1
    },
  },
  LEFT: { args: 's i?', fn: (s: string, n = 1) => (n < 0 ? E.VALUE : s.substring(0, n)) },
  LEN: { args: 's', fn: (s: string) => s.length },
  LOWER: { args: 's', fn: (s: string) => s.toLowerCase() },
  MID: {
    args: 's i i',
    fn: (s: string, start: number, n: number) => (start < 1 || start > MAX_POS || n < 0 ? E.VALUE : s.substring(start - 1, start - 1 + n)),
  },
  PROPER: { args: 's', fn: proper },
  REPLACE: {
    args: 's i i s',
    fn: (s: string, start: number, n: number, rep: string) => (start < 1 || start > MAX_POS || n < 0 ? E.VALUE : capped(s.slice(0, start - 1) + rep + s.slice(start - 1 + n))),
  },
  REPT: { args: 's i', fn: (s: string, n: number) => (n < 0 || s.length * n > MAX_TEXT ? E.VALUE : s.repeat(n)) },
  RIGHT: { args: 's i?', fn: (s: string, n = 1) => (n < 0 ? E.VALUE : s.substring(Math.max(0, s.length - n))) },
  SEARCH: { args: 's s i?', fn: search },
  SUBSTITUTE: { args: 's s s i?', fn: substitute },
  T: { args: 'a', fn: (v: V) => (typeof v === 'string' || isErr(v) ? v : '') },
  TEXTJOIN: { args: 'V b V', rep: 1, fn: textjoin },
  // Excel trims spaces only, not tabs or line breaks
  TRIM: { args: 's', fn: (s: string) => s.replace(/ +/g, ' ').replace(/^ | $/g, '') },
  UNICHAR: { args: 'i', fn: (n: number) => (n < 1 || n > 0x10ffff ? E.VALUE : n >= 0xd800 && n <= 0xdfff ? E.NA : String.fromCodePoint(n)) },
  UNICODE: { args: 's', fn: (s: string) => (s === '' ? E.VALUE : s.codePointAt(0)!) },
  UPPER: { args: 's', fn: upper },
}
