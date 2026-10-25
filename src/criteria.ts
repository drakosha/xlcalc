/**
 * Criteria of SUMIF, SUMIFS, COUNTIFS and an index over them: strings are case-insensitive but
 * accent-sensitive («й» ≠ «и», «ё» ≠ «е»); wildcards `*`, `?`, escaping with `~`; numbers, dates and times in
 * the criterion text are compared as numbers, and a number criterion's `=` also holds for text that reads as
 * that number (as VALUE reads it), while its `<>` holds for all text; `""` matches blanks and empty text,
 * `"="` blanks alone, `"<>"` the rest; an error, or its text, matches that error. Numbers, dates and words in the
 * text and the order of text follow the book's locale (`CritRules`, en-US by default).
 */
import type { ErrorCode } from './contract.js'
import { collator } from './collate.js'
import { critTextToNumber, textToNumber } from './dates.js'
import { spend } from './exec/budget.js'
import { ERROR_TYPE } from './error-type.js'
import { readsClock, type RulesClock } from './locale/rules/data.js'
import { CELL_ERRORS, ERROR_CODE } from './syntax/errors.js'
import { holds, to15 } from './values.js'
import { wildMatcher, type Wildcard } from './wildcards.js'

/** A cell value for a criterion; errors are any object. */
export type CritValue = number | string | boolean | null | object

export type Crit =
  | { op: string; kind: 'num'; n: number }
  | { op: string; kind: 'bool'; b: boolean }
  | { op: string; kind: 'str'; s: string }
  | { op: string; kind: 'wild'; re: Wildcard }
  | { op: string; kind: 'blank' }
  | { op: string; kind: 'err'; code: string }

/** A criterion's error text by a locale's error texts (`'#N/A': '#Н/Д'`; one not given keeps its English text): the error code, or undefined. */
export function errorWords(local: Readonly<Record<string, string>>): (text: string) => string | undefined {
  const words = new Map(CELL_ERRORS.map((t) => [(local[t] ?? t).toUpperCase(), ERROR_CODE[t]]))
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
      c = re ? { op, kind: 'wild', re } : { op, kind: 'str', s: rest.toLowerCase() }
    }
  }
  return c
}

const ord = (x: number, y: number) => (x < y ? -1 : x > y ? 1 : 0)

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
      // errors only, never text that reads like one; < and > order errors as ERROR.TYPE does (#ERROR! counts as #NULL!)
      const code = v !== null && typeof v === 'object' ? (v as { code?: string }).code : undefined
      if (c.op === '=' || c.op === '<>') return (code === c.code) === (c.op === '=')
      const r = code === undefined ? undefined : ERROR_TYPE[code as ErrorCode]
      return r !== undefined && holds(c.op, ord(r, ERROR_TYPE[c.code as ErrorCode] as number))
    }
    case 'str':
    case 'low':
      if (typeof v !== 'string') return c.op === '<>'
      return textHolds(c, c.kind === 'str' ? v.toLowerCase() : v, rules)
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
    if (low === undefined) seen.set(v, (low = v.toLowerCase()))
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
