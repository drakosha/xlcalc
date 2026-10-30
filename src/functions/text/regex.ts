/**
 * REGEXTEST, REGEXEXTRACT and REGEXREPLACE over the level of regular expressions in use (`regex-level.ts`), the one of
 * XLOOKUP's match mode 3, by Microsoft's description: case-insensitive with `case_sensitivity` 1 (the pattern under a
 * leading `(?i)`), every match taken from the end of the one before it (an empty one a character on), `$n` and `${n}`
 * in a replacement for a group, `$$` for a dollar. A pattern that is none and Excel's match limit are #VALUE!, and so
 * is an answer longer than 32,767 characters. Neither Excel the harness drives has these functions (Excel for Mac
 * 16.114, Excel for Windows 16.0.20430: #NAME?), so they are not checked against Excel.
 */
import type { FnDef } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { spend } from '../../exec/budget.js'
import { regexPattern } from '../../regex-level.js'
import { E, type CErr, type V } from '../../values.js'
import { capped } from './common.js'

/** The pattern compiled, case-insensitive for 1, or the error of a case sensitivity other than 0 and 1 or of no pattern. */
function compiled(pattern: string, caseless: number) {
  if (caseless !== 0 && caseless !== 1) return E.VALUE
  return regexPattern(caseless ? `(?i)${pattern}` : pattern) ?? E.VALUE
}

/** Every match of `re` in `s` in order (each its start, end and groups' spans), or `'limit'`. */
function matches(re: Exclude<ReturnType<typeof compiled>, CErr>, s: string): (readonly number[])[] | 'limit' {
  const out: (readonly number[])[] = []
  for (let at = 0; at <= s.length; ) {
    spend(1)
    const m = re.exec(s, at)
    if (m === 'limit') return m
    if (!m) break
    out.push(m)
    // after an empty match the next is looked for a character on, a surrogate pair as one
    at = m[1] > m[0] ? m[1] : m[1] + (s.codePointAt(m[1])! > 0xffff ? 2 : 1)
  }
  return out
}

/** The text of a group's span, empty for a group that took no part. */
const span = (s: string, m: readonly number[], g: number) => (m[2 * g] < 0 ? '' : s.slice(m[2 * g], m[2 * g + 1]))

/** A replacement for match `m`: `$n` and `${n}` its groups (up to the pattern's), `$$` a dollar, any other `$` as it is. */
function replaced(rep: string, s: string, m: readonly number[]): string {
  const groups = m.length / 2 - 1
  return rep.replace(/\$(?:\$|\{(\d+)\}|(\d{1,2}))/g, (all, braced?: string, bare?: string) => {
    if (all === '$$') return '$'
    const g = Number(braced ?? bare)
    // `$12` with fewer groups is `$1` and a 2
    if (bare && g > groups && bare.length === 2) return Number(bare[0]) <= groups ? span(s, m, Number(bare[0])) + bare[1] : all
    return g <= groups ? span(s, m, g) : all
  })
}

/** A row of texts, or #N/A for none. */
const row = (xs: string[]): V | Arr => (xs.length ? new Arr(1, xs.length, 1, xs.length, xs) : E.NA)

export const functions: Record<string, FnDef> = {
  REGEXTEST: {
    args: 's s w?',
    fn: (s: string, pattern: string, caseless = 0) => {
      const re = compiled(pattern, caseless)
      if (!('test' in re)) return re
      const r = re.test(s)
      return r === 'limit' ? E.VALUE : r
    },
  },
  REGEXEXTRACT: {
    args: 's s w? w?',
    array: true,
    fn: (s: string, pattern: string, mode = 0, caseless = 0) => {
      if (mode !== 0 && mode !== 1 && mode !== 2) return E.VALUE
      const re = compiled(pattern, caseless)
      if (!('test' in re)) return re
      if (mode === 1) {
        const all = matches(re, s)
        return all === 'limit' ? E.VALUE : row(all.map((m) => span(s, m, 0)))
      }
      const m = re.exec(s, 0)
      if (m === 'limit') return E.VALUE
      if (!m) return E.NA
      return mode === 0 ? span(s, m, 0) : row(Array.from({ length: m.length / 2 - 1 }, (_, g) => span(s, m, g + 1)))
    },
  },
  REGEXREPLACE: {
    args: 's s s w? w?',
    fn: (s: string, pattern: string, rep: string, occurrence = 0, caseless = 0) => {
      const re = compiled(pattern, caseless)
      if (!('test' in re)) return re
      const all = matches(re, s)
      if (all === 'limit') return E.VALUE
      // 0 every match, n the nth, -n the nth from the end
      const k = occurrence > 0 ? occurrence - 1 : all.length + occurrence
      const chosen = occurrence === 0 ? all : all[k] && k >= 0 ? [all[k]] : []
      let out = ''
      let at = 0
      for (const m of chosen) {
        out += s.slice(at, m[0]) + replaced(rep, s, m)
        at = m[1]
      }
      return capped(out + s.slice(at))
    },
  },
}
