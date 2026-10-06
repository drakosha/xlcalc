/**
 * Criteria of SUMIF, SUMIFS, COUNTIFS and an index over them: strings are case-insensitive but
 * accent-sensitive («й» ≠ «и», «ё» ≠ «е»); wildcards `*`, `?`, escaping with `~`; numbers, dates and times in
 * the criterion text are compared as numbers, and a number criterion's `=` also holds for text that reads as
 * that number (as VALUE reads it), while its `<>` holds for all text; `""` matches blanks, `"<>"` non-blanks;
 * an error, or its text, matches that error. Numbers, dates and words in the text and the order of text
 * follow the book's locale (`CritRules`, en-US by default).
 */
import { textToNumber } from './dates.js'
import { to15 } from './values.js'

/** A cell value for a criterion; errors are any object. */
export type CritValue = number | string | boolean | null | object

export type Crit =
  | { op: string; kind: 'num'; n: number }
  | { op: string; kind: 'bool'; b: boolean }
  | { op: string; kind: 'str'; s: string }
  | { op: string; kind: 'wild'; re: RegExp }
  | { op: string; kind: 'blank' }
  | { op: string; kind: 'err'; code: string }

/**
 * String order for `<`/`>`: case-insensitive, accent-sensitive. Without `Intl` (QuickJS, for example)
 * the order falls back to code units of the lower-cased strings.
 */
export const collator: { compare(a: string, b: string): number } =
  typeof Intl !== 'undefined' ? new Intl.Collator('en-US', { sensitivity: 'accent' }) : { compare: (a, b) => ((a = a.toLowerCase()) < (b = b.toLowerCase()) ? -1 : a > b ? 1 : 0) }

/** What criteria read by the locale: numbers, dates and words in a criterion's text, numbers in cell text, the order of text. */
export interface CritRules {
  /** A number or a date in a criterion's text (after the operator, trimmed), or null. */
  critNumber(text: string): number | null
  /** TRUE or FALSE in a criterion's text (after the operator), or null. */
  critBool(text: string): boolean | null
  /** Cell text that a number criterion's equality takes as that number, or null. */
  cellNumber(text: string): number | null
  /** The order of two texts for `<` and `>`, and in sorting and lookups: case-insensitive, accent-sensitive. */
  compareText(a: string, b: string): number
}

/** en-US: numbers, dates and times in criteria and in cell text as VALUE reads them, TRUE and FALSE in any case. */
export const EN_CRIT: CritRules = {
  critNumber: textToNumber,
  critBool: (t) => (/^(TRUE|FALSE)$/i.test(t) ? t.toUpperCase() === 'TRUE' : null),
  cellNumber: (s) => (/\d/.test(s) ? textToNumber(s) : null),
  compareText: collator.compare,
}

/** Parsed criteria by value, per locale's rules. */
const caches = new WeakMap<CritRules, Map<string, Crit>>()

const literal = (ch: string) => ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

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

/** Error texts in a criterion (any case) and their order, ERROR.TYPE's numbers; #ERROR! is how this engine reads #NULL!. */
const ERRORS: Record<string, [code: string, rank: number]> = {
  '#NULL!': ['ERROR', 1],
  '#DIV/0!': ['DIV0', 2],
  '#VALUE!': ['VALUE', 3],
  '#REF!': ['REF', 4],
  '#NAME?': ['NAME', 5],
  '#NUM!': ['NUM', 6],
  '#N/A': ['NA', 7],
  '#SPILL!': ['SPILL', 9],
  '#CALC!': ['CALC', 14],
}
const ERROR_RANK: Record<string, number> = Object.fromEntries(Object.values(ERRORS))

/** A criterion that is an error value: it matches the cells holding that error (Excel: COUNTIF(A1:A9,NA())). */
export const errorCrit = (code: string): Crit => ({ op: '=', kind: 'err', code })

/** A criterion from a value (the scalar after evaluating the argument). */
export function parseCrit(v: number | string | boolean | null, rules: CritRules = EN_CRIT): Crit {
  let cache = caches.get(rules)
  if (!cache) caches.set(rules, (cache = new Map()))
  const key = typeof v + ':' + String(v)
  const hit = cache.get(key)
  if (hit) return hit
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
    if (rest === '') c = { op, kind: 'blank' }
    else if (rest.toUpperCase() in ERRORS) c = { op, kind: 'err', code: ERRORS[rest.toUpperCase()][0] }
    else if ((n = rules.critNumber(rest.trim())) !== null) c = { op, kind: 'num', n }
    else if ((b = rules.critBool(rest)) !== null) c = { op, kind: 'bool', b }
    else {
      const re = op === '=' || op === '<>' ? wildRe(rest) : null
      c = re ? { op, kind: 'wild', re } : { op, kind: 'str', s: rest.toLowerCase() }
    }
  }
  cache.set(key, c)
  return c
}

const ord = (x: number, y: number) => (x < y ? -1 : x > y ? 1 : 0)
const opOk = (op: string, c: number) => (op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0)

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
      return opOk(c.op, ord(to15(v), to15(c.n)))
    case 'bool':
      if (typeof v !== 'boolean') return c.op === '<>'
      return opOk(c.op, ord(Number(v), Number(c.b)))
    case 'blank':
      if (c.op === '=') return v === null || v === ''
      if (c.op === '<>') return !(v === null || v === '')
      return false
    case 'wild':
      if (typeof v !== 'string') return c.op === '<>'
      return c.re.test(v) === (c.op === '=')
    case 'err': {
      // errors only, never text that reads like one; < and > order errors as ERROR.TYPE does
      const code = v !== null && typeof v === 'object' ? (v as { code?: string }).code : undefined
      if (c.op === '=' || c.op === '<>') return (code === c.code) === (c.op === '=')
      const r = code === undefined ? undefined : ERROR_RANK[code]
      return r !== undefined && opOk(c.op, ord(r, ERROR_RANK[c.code]))
    }
    case 'str':
      if (typeof v !== 'string') return c.op === '<>'
      if (c.op === '=' || c.op === '<>') return (v.toLowerCase() === c.s) === (c.op === '=')
      return opOk(c.op, rules.compareText(v.toLowerCase(), c.s))
  }
}

/** Equality key in the index: case-insensitive strings, numbers (and text that reads as a number), booleans. */
function eqKey(v: CritValue, rules: CritRules): string | null {
  if (typeof v === 'string') {
    const n = rules.cellNumber(v)
    return n === null ? 's' + v.toLowerCase() : 'n' + to15(n)
  }
  if (typeof v === 'number') return 'n' + to15(v)
  if (typeof v === 'boolean') return 'b' + v
  return null
}

function critKey(c: Crit): string | null {
  if (c.op !== '=') return null
  if (c.kind === 'str') return 's' + c.s
  if (c.kind === 'num') return 'n' + to15(c.n)
  if (c.kind === 'bool') return 'b' + c.b
  return null
}

export type CritIndex = Map<string, number[]>

export function buildIndex(flat: readonly CritValue[], rules: CritRules = EN_CRIT): CritIndex {
  const ix: CritIndex = new Map()
  for (let i = 0; i < flat.length; i++) {
    const k = eqKey(flat[i], rules)
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
  if (cand) for (const p of cand) check(p)
  else for (let p = 0; p < n; p++) check(p)
  return out
}
