/**
 * Criteria of SUMIF, SUMIFS, COUNTIFS and an index over them: strings are case-insensitive but
 * accent-sensitive («й» ≠ «и», «ё» ≠ «е»); wildcards `*`, `?`, escaping with `~`; numbers, dates and times in
 * the criterion text are compared as numbers, and a number criterion's `=` also holds for text that reads as
 * that number (as VALUE reads it), while its `<>` holds for all text; `""` matches blanks and empty text,
 * `"="` blanks alone, `"<>"` the rest; an error, or its text, matches that error. Numbers, dates and words in the
 * text and the order of text follow the book's locale (`CritRules`, en-US by default).
 */
import { collator } from './collate.js'
import { critTextToNumber, textToNumber } from './dates.js'
import { spend } from './exec/budget.js'
import { readsClock, type RulesClock } from './locale/rules/data.js'
import { holds, to15 } from './values.js'

/** A cell value for a criterion; errors are any object. */
export type CritValue = number | string | boolean | null | object

export type Crit =
  | { op: string; kind: 'num'; n: number }
  | { op: string; kind: 'bool'; b: boolean }
  | { op: string; kind: 'str'; s: string }
  | { op: string; kind: 'wild'; re: RegExp }
  | { op: string; kind: 'blank' }
  | { op: string; kind: 'err'; code: string }

/** Error texts in a criterion (any case) and their order, ERROR.TYPE's numbers. */
const ERRORS: Record<string, [code: string, rank: number]> = {
  '#NULL!': ['NULL', 1],
  '#DIV/0!': ['DIV0', 2],
  '#VALUE!': ['VALUE', 3],
  '#REF!': ['REF', 4],
  '#NAME?': ['NAME', 5],
  '#NUM!': ['NUM', 6],
  '#N/A': ['NA', 7],
  '#SPILL!': ['SPILL', 9],
  '#CALC!': ['CALC', 14],
}
// a formula that does not compute (#ERROR!) ranks as ERROR.TYPE numbers it
const ERROR_RANK: Record<string, number> = { ...Object.fromEntries(Object.values(ERRORS)), ERROR: 1 }

/** A criterion's error text by a locale's error texts (`'#N/A': '#Н/Д'`; one not given keeps its English text): the error code, or undefined. */
export function errorWords(local: Readonly<Record<string, string>>): (text: string) => string | undefined {
  const words = new Map(Object.entries(ERRORS).map(([text, [code]]) => [(local[text] ?? text).toUpperCase(), code]))
  return (text) => words.get(text.toUpperCase())
}

/** What criteria read by the locale: numbers, dates and words in a criterion's text, numbers in cell text, the order of text. */
export interface CritRules {
  /** A number or a date in a criterion's text (after the operator; a space before a date makes it text), or null. */
  critNumber(text: string): number | null
  /** TRUE or FALSE in a criterion's text (after the operator), or null. */
  critBool(text: string): boolean | null
  /** Cell text that a number criterion's equality takes as that number, or null. */
  cellNumber(text: string): number | null
  /** The order of two texts for `<` and `>`, and in sorting and lookups: case-insensitive, accent-sensitive. */
  compareText(a: string, b: string): number
  /** An error's text in a criterion (any case): the error code, or undefined. */
  critError(text: string): string | undefined
  /** The clock a date without a year reads: a criterion or an index that read it is not cached. */
  readonly clock?: RulesClock
}

/** en-US: numbers, dates and times in criteria and in cell text as VALUE reads them, TRUE and FALSE in any case. */
export const EN_CRIT: CritRules = {
  critNumber: critTextToNumber,
  critBool: (t) => (/^(TRUE|FALSE)$/i.test(t) ? t.toUpperCase() === 'TRUE' : null),
  cellNumber: (s) => (/\d/.test(s) ? textToNumber(s) : null),
  compareText: collator.compare,
  critError: errorWords({}),
}

/** Parsed criteria by value, per locale's rules. */
const caches = new WeakMap<CritRules, Map<string, Crit>>()

/** Text matched literally in a regular expression. */
export const literal = (ch: string) => ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** The regular expression source of a wildcard pattern as Excel reads it: `*` any text, `?` one character, `~` takes the next character as it is (a trailing `~` is dropped). */
export function wildSource(s: string): string {
  let re = ''
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (ch === '~') re += i + 1 < s.length ? literal(s[++i]) : ''
    else if (ch === '*') re += '[\\s\\S]*'
    else if (ch === '?') re += '[\\s\\S]'
    else re += literal(ch)
  }
  return re
}

/**
 * A whole-text, case-insensitive matcher, or null for plain text. In criteria only `*` or `?` make a pattern
 * (`a~~c` is compared as it is); in lookups (`tilde`) a `~` alone does too (`a~c` looks up `ac`).
 */
export function wildRe(s: string, tilde = false): RegExp | null {
  return (tilde ? /[*?~]/ : /[*?]/).test(s) ? new RegExp(`^${wildSource(s)}$`, 'iu') : null
}

/** A criterion that is an error value: it matches the cells holding that error (Excel: COUNTIF(A1:A9,NA())). */
export const errorCrit = (code: string): Crit => ({ op: '=', kind: 'err', code })

/** A criterion from a value (the scalar after evaluating the argument). */
export function parseCrit(v: number | string | boolean | null, rules: CritRules = EN_CRIT): Crit {
  let cache = caches.get(rules)
  if (!cache) caches.set(rules, (cache = new Map()))
  const key = typeof v + ':' + String(v)
  const hit = cache.get(key)
  if (hit) return hit
  const [c, clocked] = readsClock(rules.clock, () => critOf(v, rules))
  if (!clocked) cache.set(key, c)
  return c
}

function critOf(v: number | string | boolean | null, rules: CritRules): Crit {
  let c: Crit
  if (typeof v === 'number') c = { op: '=', kind: 'num', n: v }
  else if (typeof v === 'boolean') c = { op: '=', kind: 'bool', b: v }
  else if (v === null) c = { op: '=', kind: 'num', n: 0 }
  else {
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(v)!
    const op = m[1] ?? '='
    const rest = m[2]
    let n: number | null = null
    let b: boolean | null = null
    let code: string | undefined
    // `""` and `"="` differ: an operator written out is kept, so `=` takes no empty text (`matches`)
    if (rest === '') c = { op: m[1] ?? '', kind: 'blank' }
    else if ((code = rules.critError(rest)) !== undefined) c = { op, kind: 'err', code }
    else if ((n = rules.critNumber(rest)) !== null) c = { op, kind: 'num', n }
    else if ((b = rules.critBool(rest)) !== null) c = { op, kind: 'bool', b }
    else {
      const re = op === '=' || op === '<>' ? wildRe(rest) : null
      c = re ? { op, kind: 'wild', re } : { op, kind: 'str', s: rest.toLowerCase() }
    }
  }
  return c
}

const ord = (x: number, y: number) => (x < y ? -1 : x > y ? 1 : 0)

export function matches(v: CritValue, c: Crit, rules: CritRules = EN_CRIT): boolean {
  switch (c.kind) {
    case 'num':
      // equality with a number also holds for text that reads as that number (Excel: 2 matches "2"), while `<>` takes all text
      if (typeof v === 'string' && c.op === '=') {
        const n = rules.cellNumber(v)
        if (n !== null) return to15(n) === to15(c.n)
      }
      if (typeof v !== 'number') return c.op === '<>'
      // numbers equal to 15 significant digits are equal, as with the comparison operators
      return holds(c.op, ord(to15(v), to15(c.n)))
    case 'bool':
      if (typeof v !== 'boolean') return c.op === '<>'
      return holds(c.op, ord(Number(v), Number(c.b)))
    case 'blank':
      // `""` takes a formula's empty text for a blank, `=` and `<>` only an empty cell, as Excel does
      if (c.op === '') return v === null || v === ''
      if (c.op === '=') return v === null
      if (c.op === '<>') return v !== null
      return false
    case 'wild':
      if (typeof v !== 'string') return c.op === '<>'
      return c.re.test(v) === (c.op === '=')
    case 'err': {
      // errors only, never text that reads like one; < and > order errors as ERROR.TYPE does
      const code = v !== null && typeof v === 'object' ? (v as { code?: string }).code : undefined
      if (c.op === '=' || c.op === '<>') return (code === c.code) === (c.op === '=')
      const r = code === undefined ? undefined : ERROR_RANK[code]
      return r !== undefined && holds(c.op, ord(r, ERROR_RANK[c.code]))
    }
    case 'str':
      if (typeof v !== 'string') return c.op === '<>'
      if (c.op === '=' || c.op === '<>') return (v.toLowerCase() === c.s) === (c.op === '=')
      return holds(c.op, rules.compareText(v.toLowerCase(), c.s))
  }
}

/** The index's key of what equality matches: text in lower case, a number to 15 digits, a boolean. */
const equalityKey = (x: string | number | boolean): string => (typeof x === 'string' ? 's' + x : typeof x === 'number' ? 'n' + to15(x) : 'b' + x)

/** Equality key of a value in the index: case-insensitive strings, numbers (and text that reads as a number), booleans. */
function eqKey(v: CritValue, rules: CritRules): string | null {
  if (typeof v === 'string') {
    const n = rules.cellNumber(v)
    return equalityKey(n === null ? v.toLowerCase() : n)
  }
  return typeof v === 'number' || typeof v === 'boolean' ? equalityKey(v) : null
}

/**
 * Equality key of a criterion, the key of the values it selects; null for one that is no equality, and for text with
 * `€`, which en-US reads as a number in a cell (its key) but not in a criterion, so such a criterion scans.
 */
export function critKey(c: Crit): string | null {
  if (c.op !== '=') return null
  return c.kind === 'str' ? (c.s.includes('€') ? null : equalityKey(c.s)) : c.kind === 'num' ? equalityKey(c.n) : c.kind === 'bool' ? equalityKey(c.b) : null
}

/** One-column COUNTIF uses the same equality keys as the selection index; other criteria use the scan rules. */
export function criterionMatcher(c: Crit, rules: CritRules): (v: CritValue) => boolean {
  const key = critKey(c)
  if (key === null) return (v) => matches(v, c, rules)
  // Per read only: text dates without a year must read the clock again on the next call.
  const texts = new Map<string, string>()
  return (v) => {
    if (typeof v !== 'string') return eqKey(v, rules) === key
    let value = texts.get(v)
    if (value === undefined) { value = eqKey(v, rules)!; texts.set(v, value) }
    return value === key
  }
}

export type CritIndex = Map<string, number[]>

/** Values between two checks of the computation's budget in long loops (65,536 − 1, a mask). */
const STRIDE = 0xffff

export function buildIndex(flat: readonly CritValue[], rules: CritRules = EN_CRIT): CritIndex {
  const ix: CritIndex = new Map()
  // a text repeated down a column is read once (a text always has a key)
  const texts = new Map<string, string>()
  const textKey = (t: string) => {
    const k = eqKey(t, rules)!
    texts.set(t, k)
    return k
  }
  for (let i = 0; i < flat.length; i++) {
    if ((i & STRIDE) === 0) spend(Math.min(STRIDE + 1, flat.length - i))
    const v = flat[i]
    const k = typeof v === 'string' ? (texts.get(v) ?? textKey(v)) : eqKey(v, rules)
    if (k === null) continue
    const list = ix.get(k)
    if (list) list.push(i)
    else ix.set(k, [i])
  }
  return ix
}

export interface CritColumn {
  /** Values by position; positions past the end are blank. */
  flat: readonly CritValue[]
  crit: Crit
  /** The index of this range (built lazily and cached by the caller). */
  index: () => CritIndex
}

/**
 * Positions where all criteria hold. The first equality criterion goes through the index, the rest are
 * checked on its candidates; without an equality criterion it is a full scan.
 */
export function select(cols: CritColumn[], n: number, rules: CritRules = EN_CRIT): number[] {
  let cand: number[] | null = null
  let used = -1
  for (let i = 0; i < cols.length; i++) {
    const k = critKey(cols[i].crit)
    if (k === null) continue
    cand = cols[i].index().get(k) ?? []
    used = i
    break
  }
  const out: number[] = []
  const check = (p: number) => {
    for (let i = 0; i < cols.length; i++) if (i !== used && !matches(cols[i].flat[p] ?? null, cols[i].crit, rules)) return
    out.push(p)
  }
  if (cand) {
    spend(cand.length)
    for (const p of cand) check(p)
  } else
    for (let p = 0; p < n; p++) {
      if ((p & STRIDE) === 0) spend(Math.min(STRIDE + 1, n - p))
      check(p)
    }
  return out
}
