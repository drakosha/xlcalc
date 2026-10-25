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
  | { op: string; kind: 'wild'; re: Wildcard }
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

/** Parsed criteria by value (a Map key: `1` and `'1'` apart, `-0` with `0`), per locale's rules. */
const caches = new WeakMap<CritRules, Map<number | string | boolean | null, Crit>>()

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
 * A wildcard pattern by its parts between `*`, each a regular expression without `*` (`?` one character, `~` the next
 * one as it is): one expression with a `[\s\S]*` for each `*` backtracks over the text once per `*` (three of them on
 * 8,000 characters run a minute in one call no deadline stops), while parts found leftmost in turn cost a search each.
 * Leftmost is enough: a later place of a part leaves less text for the parts after it. `flags` are the expressions'
 * own (`iu` in criteria, lookups and SEARCH; none in Find, which folds the text itself).
 */
export class Wildcard {
  /** The whole pattern as one expression anchored at both ends, with one `*` at most (it backtracks once, over a part's length); else null. */
  private readonly whole: RegExp | null
  /** The first part at the start (sticky), the parts found anywhere, and the last part at the end; null for an empty part. */
  private readonly head: RegExp | null
  private readonly parts: (RegExp | null)[]
  private readonly tail: RegExp | null

  /** Nothing to match: an empty pattern, or `~` alone. */
  readonly empty: boolean

  constructor(pattern: string, flags: string) {
    const sources = ['']
    for (let i = 0; i < pattern.length; i++) {
      const ch = pattern[i]
      if (ch === '*') sources.push('')
      else sources[sources.length - 1] += ch === '~' ? (i + 1 < pattern.length ? literal(pattern[++i]) : '') : ch === '?' ? '[\\s\\S]' : literal(ch)
    }
    const k = sources.length - 1
    this.empty = !k && !sources[0]
    const re = (source: string, more: string, after = '') => (source ? new RegExp(`(?:${source})${after}`, flags + more) : null)
    this.whole = k > 1 ? null : new RegExp(`^(?:${sources.join('[\\s\\S]*')})$`, flags)
    this.head = k > 1 ? re(sources[0], 'y') : null
    this.parts = sources.map((x) => re(x, 'g'))
    this.tail = k > 1 ? re(sources[k], 'g', '$') : null
  }

  /** Whether the whole text matches. */
  test(text: string): boolean {
    if (this.whole) return this.whole.test(text)
    let p = 0
    if (this.head) {
      this.head.lastIndex = 0
      if (!this.head.test(text)) return false
      p = this.head.lastIndex
    }
    p = this.middle(text, p)
    if (p < 0 || !this.tail) return p >= 0
    this.tail.lastIndex = p
    return this.tail.test(text)
  }

  /**
   * The first match at or after `from`, as [start, end], or null: each `*` takes as little as it can, a last one the
   * rest of the text (Replace in Find), and the start is the leftmost one there is (SEARCH).
   */
  find(text: string, from = 0): [start: number, end: number] | null {
    const first = this.parts[0]
    let start = from
    let p = from
    if (first) {
      first.lastIndex = from
      const m = first.exec(text)
      if (!m) return null
      start = m.index
      p = first.lastIndex
    }
    if (this.parts.length === 1) return [start, p]
    p = this.middle(text, p)
    if (p < 0) return null
    const last = this.parts[this.parts.length - 1]
    if (!last) return [start, text.length]
    last.lastIndex = p
    return last.test(text) ? [start, last.lastIndex] : null
  }

  /** Where the parts between the first and the last end, found leftmost in turn from `p`; −1 where one is not found. */
  private middle(text: string, p: number): number {
    for (let k = 1; k < this.parts.length - 1; k++) {
      const re = this.parts[k]
      if (!re) continue
      re.lastIndex = p
      if (!re.test(text)) return -1
      p = re.lastIndex
    }
    return p
  }
}

/**
 * A whole-text, case-insensitive matcher, or null for plain text. In criteria only `*` or `?` make a pattern
 * (`a~~c` is compared as it is); in lookups (`tilde`) a `~` alone does too (`a~c` looks up `ac`).
 */
export function wildMatcher(s: string, tilde = false): Wildcard | null {
  return (tilde ? /[*?~]/ : /[*?]/).test(s) ? new Wildcard(s, 'iu') : null
}

/** Criteria of error values by code, one each, so that a criterion is found by itself (kept selections). */
const errorCrits = new Map<string, Crit>()

/** A criterion that is an error value: it matches the cells holding that error (Excel: COUNTIF(A1:A9,NA())). */
export function errorCrit(code: string): Crit {
  let c = errorCrits.get(code)
  if (!c) errorCrits.set(code, (c = { op: '=', kind: 'err', code }))
  return c
}

/** A criterion from a value (the scalar after evaluating the argument). */
export function parseCrit(v: number | string | boolean | null, rules: CritRules = EN_CRIT): Crit {
  let cache = caches.get(rules)
  if (!cache) caches.set(rules, (cache = new Map()))
  const hit = cache.get(v)
  if (hit) return hit
  const [c, clocked] = readsClock(rules.clock, () => critOf(v, rules))
  if (!clocked) cache.set(v, c)
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
      const re = op === '=' || op === '<>' ? wildMatcher(rest) : null
      c = re ? { op, kind: 'wild', re } : { op, kind: 'str', s: lowered(rest) }
    }
  }
  return c
}

const ord = (x: number, y: number) => (x < y ? -1 : x > y ? 1 : 0)

/** Text as a criterion's equality takes it: in lower case, é as one code point or as e and an accent alike, as Excel's COUNTIF counts them (the set `low-values-held`). */
const lowered = (s: string) => (NON_ASCII.test(s) ? s.normalize('NFC') : s).toLowerCase()
const NON_ASCII = /[^\x00-\x7f]/

/** A text criterion over texts already in lower case: a range's kept ones (`select`). */
interface Lowered {
  op: string
  kind: 'low'
  s: string
}

export function matches(v: CritValue, c: Crit | Lowered, rules: CritRules = EN_CRIT): boolean {
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
    case 'low':
      if (typeof v !== 'string') return c.op === '<>'
      return textHolds(c, c.kind === 'str' ? lowered(v) : v, rules)
  }
}

/** Whether a text criterion holds for a text in lower case. */
function textHolds(c: { op: string; s: string }, low: string, rules: CritRules): boolean {
  if (c.op === '=' || c.op === '<>') return (low === c.s) === (c.op === '=')
  return holds(c.op, rules.compareText(low, c.s))
}

/** A range's texts in lower case by position; undefined where it holds no text. */
export type Lower = readonly (string | undefined)[]

/** The texts of `flat` in lower case, each distinct text lowered once, and how many distinct texts it holds. */
export function lowerOf(flat: readonly CritValue[]): [Lower, number] {
  const out = new Array<string | undefined>(flat.length)
  const seen = new Map<string, string>()
  for (let i = 0; i < flat.length; i++) {
    if ((i & STRIDE) === 0) spend(Math.min(STRIDE + 1, flat.length - i))
    const v = flat[i]
    if (typeof v !== 'string') continue
    let low = seen.get(v)
    if (low === undefined) seen.set(v, (low = lowered(v)))
    out[i] = low
  }
  return [out, seen.size]
}

/** The index's key of what equality matches: text in lower case, a number to 15 digits, a boolean. */
const equalityKey = (x: string | number | boolean): string => (typeof x === 'string' ? 's' + x : typeof x === 'number' ? 'n' + to15(x) : 'b' + x)

/** Equality key of a value in the index: case-insensitive strings, numbers (and text that reads as a number), booleans. */
function eqKey(v: CritValue, rules: CritRules): string | null {
  if (typeof v === 'string') {
    const n = rules.cellNumber(v)
    return equalityKey(n === null ? lowered(v) : n)
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

/** What tests count of selections (`watchSelections`): the selections made and the candidates they checked. Nothing is counted unless a test asks. */
export interface SelectWatch {
  selects: number
  checks: number
}
let watch: SelectWatch | undefined

/** Counts the selections made from now on into `w`; undefined stops. For tests. */
export function watchSelections(w?: SelectWatch): void {
  watch = w
}

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
  /** Its texts in lower case, where the caller keeps them (from the range's second read on). */
  lower?: () => Lower | undefined
  /** Its index where the caller keeps one, else, with `build`, where it read the range before (then built and kept). */
  ready?: (build: boolean) => CritIndex | undefined
}

/** Positions of a criterion past which the indexes of the other equality criteria are worth looking at. */
const WIDE = 64

/**
 * Positions where all criteria hold. The first equality criterion goes through the index, the rest are checked on its
 * candidates; where it gives `WIDE` positions or more, another equality criterion's index at hand that gives fewer does,
 * one built for it where the first gives a sixteenth of the positions or more; without an equality criterion it is a
 * full scan. The positions come in order either way.
 */
export function select(cols: CritColumn[], n: number, rules: CritRules = EN_CRIT): number[] {
  if (watch) watch.selects++
  let cand: number[] | null = null
  let used = -1
  for (let i = 0; i < cols.length; i++) {
    const k = critKey(cols[i].crit)
    if (k === null) continue
    if (cand && cand.length < WIDE) break
    const ix: CritIndex | undefined = cand ? cols[i].ready?.(cand.length * 16 >= n) : cols[i].index()
    const list: number[] | null = ix ? (ix.get(k) ?? []) : null
    if (list && (!cand || list.length < cand.length)) {
      cand = list
      used = i
    }
  }
  // a text criterion checks a range's texts kept in lower case, where they are kept
  const flats: (readonly CritValue[] | Lower)[] = []
  const crits: (Crit | Lowered)[] = []
  for (let i = 0; i < cols.length; i++) {
    const { flat, crit } = cols[i]
    const low = i !== used && crit.kind === 'str' ? cols[i].lower?.() : undefined
    flats.push(low ?? flat)
    crits.push(low ? { op: crit.op, kind: 'low', s: (crit as Extract<Crit, { kind: 'str' }>).s } : crit)
  }
  const out: number[] = []
  const check = (p: number) => {
    for (let i = 0; i < cols.length; i++) if (i !== used && !matches(flats[i][p] ?? null, crits[i], rules)) return
    out.push(p)
  }
  if (watch) watch.checks += cand ? cand.length : n
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
