/**
 * The matcher of match mode 3's regular expressions (`XLOOKUP`, `XMATCH`): JavaScript's own over the pattern as
 * `rewrite` gives it, or the level of `xlcalc/regex` that `registerRegex` loaded (the full translation, or the PCRE2
 * engine). A pool's threads load the same level (`worker/locales.ts`).
 */
import type { RegexLevel } from './contract.js'
import { rewrite } from './functions/lookup/regex-syntax.js'
import { InputError } from './input-error.js'

/** A pattern's test of a text, as `RegexLevel.compile` gives it. */
export type RegexTest = (text: string) => boolean | 'limit'

export const REGEX_KINDS: readonly RegexLevel['kind'][] = ['xlcalc-pcre-syntax', 'xlcalc-pcre-engine']

let level: RegexLevel | null = null
const listeners = new Set<() => void>()
const cache = new Map<string, RegexTest | null>()
const CACHED = 64

/** The registered level, or null. */
export const regexLevel = (): RegexLevel | null => level

/** Loads a level for match mode 3 and tells the listeners (a pool loads it in its threads). */
export function setRegexLevel(l: RegexLevel): void {
  if (level === l) return
  level = l
  cache.clear()
  for (const f of listeners) f()
}

/**
 * Loads a level of `xlcalc/regex` for match mode 3 of `XLOOKUP` and `XMATCH` in the calculations from now on, a pool's
 * threads too: `import { pcreEngine } from 'xlcalc/regex'`, then `registerRegex(pcreEngine)` (or `pcreSyntax`). The
 * last one registered is used.
 */
export function registerRegex(l: RegexLevel): void {
  if (typeof l !== 'object' || l === null || !REGEX_KINDS.includes(l.kind) || typeof l.compile !== 'function') throw new InputError('level', "expected pcreSyntax or pcreEngine of 'xlcalc/regex'")
  setRegexLevel(l)
}

/** Hears a level being loaded; the answer stops it. */
export function onRegexLevel(f: () => void): () => void {
  listeners.add(f)
  return () => listeners.delete(f)
}

/** A pattern's test, or null for one that is no pattern: `#VALUE!`. */
export function regexTest(pattern: string): RegexTest | null {
  let t = cache.get(pattern)
  if (t === undefined) {
    const p = level ? level.compile(pattern) : rewrite(pattern)
    t = p && ((s) => p.test(s))
    if (cache.size >= CACHED) cache.clear()
    cache.set(pattern, t)
  }
  return t
}
