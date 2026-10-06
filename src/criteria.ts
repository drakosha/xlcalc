/**
 * Criteria of SUMIF, SUMIFS, COUNTIFS and an index over them: strings are case-insensitive but
 * accent-sensitive («й» ≠ «и», «ё» ≠ «е»); wildcards `*`, `?`, escaping with `~`; numbers and ISO dates in
 * the criterion text are compared as numbers, and a number criterion's equality also holds for text that
 * reads as that number; `""` matches blanks, `"<>"` non-blanks; an error, or its text, matches that error.
 */
import { dateToSerial, NUMBER_TEXT, readNumber } from './model.js'

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
  typeof Intl !== 'undefined' ? new Intl.Collator('ru-RU', { sensitivity: 'accent' }) : { compare: (a, b) => ((a = a.toLowerCase()) < (b = b.toLowerCase()) ? -1 : a > b ? 1 : 0) }
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/
const cache = new Map<string, Crit>()

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
}
const ERROR_RANK: Record<string, number> = Object.fromEntries(Object.values(ERRORS))

/** A criterion that is an error value: it matches the cells holding that error (Excel: COUNTIF(A1:A9,NA())). */
export const errorCrit = (code: string): Crit => ({ op: '=', kind: 'err', code })

/** A criterion from a value (the scalar after evaluating the argument). */
export function parseCrit(v: number | string | boolean | null): Crit {
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
    const t = rest.trim()
    if (rest === '') c = { op, kind: 'blank' }
    else if (rest.toUpperCase() in ERRORS) c = { op, kind: 'err', code: ERRORS[rest.toUpperCase()][0] }
    else if (NUMBER_TEXT.test(t)) c = { op, kind: 'num', n: readNumber(t) }
    else if (DATE_RE.test(t)) c = { op, kind: 'num', n: dateToSerial(t) }
    else if (/^(TRUE|FALSE)$/i.test(rest)) c = { op, kind: 'bool', b: rest.toUpperCase() === 'TRUE' }
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

/** Text that reads as a number, as in a criterion (`"2"`, `" 2.50 "`), or null. */
function numericText(s: string): number | null {
  const t = s.trim()
  return NUMBER_TEXT.test(t) ? readNumber(t) : null
}

export function matches(v: CritValue, c: Crit): boolean {
  switch (c.kind) {
    case 'num':
      // equality with a number also holds for text that reads as that number (Excel: 2 matches "2")
      if (typeof v === 'string' && (c.op === '=' || c.op === '<>')) {
        const n = numericText(v)
        if (n !== null) return (n === c.n) === (c.op === '=')
      }
      if (typeof v !== 'number') return c.op === '<>'
      return opOk(c.op, ord(v, c.n))
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
      return opOk(c.op, collator.compare(v.toLowerCase(), c.s))
  }
}

/** Equality key in the index: case-insensitive strings, numbers (and text that reads as a number), booleans. */
function eqKey(v: CritValue): string | null {
  if (typeof v === 'string') {
    const n = numericText(v)
    return n === null ? 's' + v.toLowerCase() : 'n' + n
  }
  if (typeof v === 'number') return 'n' + v
  if (typeof v === 'boolean') return 'b' + v
  return null
}

function critKey(c: Crit): string | null {
  if (c.op !== '=') return null
  if (c.kind === 'str') return 's' + c.s
  if (c.kind === 'num') return 'n' + c.n
  if (c.kind === 'bool') return 'b' + c.b
  return null
}

export type CritIndex = Map<string, number[]>

export function buildIndex(flat: readonly CritValue[]): CritIndex {
  const ix: CritIndex = new Map()
  for (let i = 0; i < flat.length; i++) {
    const k = eqKey(flat[i])
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
export function select(cols: CritColumn[], n: number): number[] {
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
    for (let i = 0; i < cols.length; i++) if (i !== used && !matches(cols[i].flat[p] ?? null, cols[i].crit)) return
    out.push(p)
  }
  if (cand) for (const p of cand) check(p)
  else for (let p = 0; p < n; p++) check(p)
  return out
}
