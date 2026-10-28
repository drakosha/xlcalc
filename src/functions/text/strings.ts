/**
 * Text functions on strings. Ported from formula.js 4.6.1 (src/text.js), Copyright (c) 2014 Sutoiku, Inc.,
 * MIT License (THIRD_PARTY.md); fixed against Excel en-US, see the comments.
 */
import { directError, eachValue, type Cx, type FnDef, type Segment } from '../../args.js'
import { Wildcard } from '../../wildcards.js'
import { remembered } from '../../memo.js'
import { E, isErr, type V } from '../../values.js'
import { NEWER_MAX_ARGS } from '../common.js'
import { capped, delimiters, joined, MAX_TEXT } from './common.js'

/** The largest position or count MID accepts; REPLACE takes up to Excel's text length. */
const MAX_INT = 2_147_483_647

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

/**
 * Excel capitalizes every letter that follows a non-letter, in any alphabet ("2-way" → "2-Way"). A run of letters is
 * lowered whole when that keeps its length (no lowering is shorter, so each letter kept its own) and it has no Σ,
 * whose whole-text lowering at a word's end is ς; else letter by letter.
 */
function proper(s: string): string {
  return s.replace(/\p{L}+/gu, (run) => {
    const n = run.codePointAt(0)! > 0xffff ? 2 : 1
    const head = caseOf(run.slice(0, n), true)
    const rest = run.slice(n)
    const low = rest.toLowerCase()
    if (low.length === rest.length && !rest.includes('Σ')) return head + low
    let out = head
    for (const ch of rest) out += caseOf(ch, false)
    return out
  })
}

// A pattern filled down a column is compiled once; the matchers keep no state between finds.
const patterns = new Map<string, Wildcard>()

/** Wildcards as in the criteria (`src/criteria.ts`); a `~` alone is an escape too. */
function search(find: string, within: string, start = 1): V {
  if (start < 1 || start > within.length) return E.VALUE
  if (find === '') return start
  const w = remembered(patterns, find, 256, () => new Wildcard(find, 'iu'))
  // the slice, not a start inside the text: under `u` a start inside a surrogate pair would step back to its lead
  const m = w.find(within.slice(start - 1))
  return m ? m[0] + start : E.VALUE
}

function substitute(text: string, old: string, rep: string, instance?: number): V {
  if (instance !== undefined && instance < 1) return E.VALUE
  if (old === '') return text
  if (instance === undefined) return capped(text.split(old).join(rep))
  // Excel counts overlapping occurrences: the 2nd "aa" in "aaa" starts at the 2nd "a"
  let at = -1
  for (let k = 0; k < instance; k++) {
    at = text.indexOf(old, at + 1)
    if (at < 0) return text
  }
  return capped(text.slice(0, at) + rep + text.slice(at + old.length))
}

/** Ranges are read row by row; several delimiters are used in turn. */
function textjoin(cx: Cx, delims: Segment, ignoreEmpty: boolean, segs: Segment[]): V {
  const ds = delimiters(delims, cx.rules)
  if (isErr(ds)) return ds
  const parts: string[] = []
  const e = directError(segs) ?? eachValue(segs, (v) => {
    const t = cx.rules.toText(v)
    if (isErr(t)) return t
    if (!(ignoreEmpty && t === '')) parts.push(t)
  })
  if (e) return e
  let out = parts[0] ?? ''
  for (let i = 1; i < parts.length; i++) out += ds[(i - 1) % ds.length] + parts[i]
  return joined(out)
}

const strings: Record<string, FnDef> = {
  // Excel removes only codes 0-31
  CLEAN: { args: 's', fn: (s: string) => s.replace(/[\0-\x1F]/g, '') },
  CONCAT: {
    args: 'V',
    rep: 1,
    max: NEWER_MAX_ARGS,
    cx: true,
    fn: (cx: Cx, segs: Segment[]) => {
      let out = ''
      const e = directError(segs) ?? eachValue(segs, (v) => {
        const t = cx.rules.toText(v)
        if (isErr(t)) return t
        out += t
      })
      return e ?? joined(out)
    },
  },
  // like &, Excel cuts the text at 32,767 characters (CONCAT and TEXTJOIN refuse it)
  CONCATENATE: { args: 's', rep: 1, fn: (ts: string[]) => ts.join('').slice(0, MAX_TEXT) },
  EXACT: { args: 's s', fn: (a: string, b: string) => a === b },
  FIND: {
    args: 's s w?',
    fn: (find: string, within: string, start = 1) => {
      if (start < 1 || start > within.length) return E.VALUE
      const at = within.indexOf(find, start - 1)
      return at < 0 ? E.VALUE : at + 1
    },
  },
  LEFT: { args: 's w?', fn: (s: string, n = 1) => (n < 0 ? E.VALUE : s.substring(0, n)) },
  LEN: { args: 's', fn: (s: string) => s.length },
  LOWER: { args: 's', fn: (s: string) => s.toLowerCase() },
  MID: {
    args: 's i i',
    fn: (s: string, start: number, n: number) => (start < 1 || start > MAX_INT || n < 0 || n > MAX_INT ? E.VALUE : s.substring(start - 1, start - 1 + n)),
  },
  PROPER: { args: 's', fn: proper },
  REPLACE: {
    args: 's w w s',
    fn: (s: string, start: number, n: number, rep: string) => (start < 1 || start > MAX_TEXT || n < 0 || n > MAX_TEXT ? E.VALUE : capped(s.slice(0, start - 1) + rep + s.slice(start - 1 + n))),
  },
  REPT: { args: 's i', fn: (s: string, n: number) => (n < 0 || s.length * n > MAX_TEXT ? E.VALUE : s.repeat(n)) },
  RIGHT: { args: 's w?', fn: (s: string, n = 1) => (n < 0 ? E.VALUE : s.substring(Math.max(0, s.length - n))) },
  SEARCH: { args: 's s w?', fn: search },
  SUBSTITUTE: { args: 's s s w?', fn: substitute },
  // a range of several cells gives its top-left cell (an array is lifted)
  T: { args: 'a', ranges: 'first', fn: (v: V) => (typeof v === 'string' || isErr(v) ? v : '') },
  TEXTJOIN: { args: 'V b V', rep: 1, max: NEWER_MAX_ARGS, cx: true, fn: textjoin },
  // Excel trims spaces only, not tabs or line breaks
  TRIM: { args: 's', fn: (s: string) => s.replace(/ +/g, ' ').replace(/^ | $/g, '') },
  // surrogates and noncharacters (U+FDD0…U+FDEF, U+xFFFE, U+xFFFF) are #N/A in Excel
  UNICHAR: { args: 'i', fn: (n: number) => (n < 1 || n > 0x10ffff ? E.VALUE : (n >= 0xd800 && n <= 0xdfff) || (n >= 0xfdd0 && n <= 0xfdef) || (n & 0xfffe) === 0xfffe ? E.NA : String.fromCodePoint(n)) },
  UNICODE: { args: 's', fn: (s: string) => (s === '' ? E.VALUE : s.codePointAt(0)!) },
  UPPER: { args: 's', fn: upper },
}

/** The byte-counting variants count characters, as in Excel without a double-byte language (en-US). */
export const functions: Record<string, FnDef> = {
  ...strings,
  FINDB: strings.FIND,
  LEFTB: strings.LEFT,
  LENB: strings.LEN,
  MIDB: strings.MID,
  REPLACEB: strings.REPLACE,
  RIGHTB: strings.RIGHT,
  SEARCHB: strings.SEARCH,
}
