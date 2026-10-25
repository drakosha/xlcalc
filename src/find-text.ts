/**
 * Excel's Find and Replace over a text, without the engine (`session/find.ts` and `session/replace.ts` over a session's
 * cells, `findMatcher` and `findReplacer` of `xlcalc/text` over a host's own texts), measured in Excel for Windows
 * (`bench/coverage/find/`): the query as it is or with Find's wildcards, a part or the whole of the text, in any case or
 * in its own.
 */
import type { FindCellsOptions } from './contract.js'
import { Wildcard } from './wildcards.js'

/** The text in lower case, each character at its place: `İ` stays (its lower case is two characters). */
export const lowerInPlace = (t: string) => {
  const lower = t.toLowerCase()
  return lower.length === t.length ? lower : t.replace(/[^İ]+/g, (m) => m.toLowerCase())
}

/**
 * Each character's lower case as one UTF-16 unit or two as it was, so `?` keeps its place (`lowerInPlace`), and the
 * final `ς` is `σ`, as `Σ` is (Find takes `σ`, `ς` and `Σ` as one letter without Match case).
 */
export const folded = (t: string) => {
  const f = lowerInPlace(t)
  return f.includes('ς') ? f.replaceAll('ς', 'σ') : f
}

/**
 * The test of a cell's text (folded without Match case), null when nothing can match (an empty query, or `~` alone).
 * Wildcards as the lookup functions read them (`wildSource`) but `?` is one UTF-16 unit, as in Find: a character
 * outside the BMP is two.
 */
export function matcher(query: string, o: FindCellsOptions): ((text: string) => boolean) | null {
  const q = o.caseSensitive ? query : folded(query)
  if (!o.wildcards) return q ? (t) => (o.wholeCell ? t === q : t.includes(q)) : null
  const w = new Wildcard(q, '')
  if (w.empty) return null
  return o.wholeCell ? (t) => t !== '' && w.test(t) : (t) => t !== '' && w.find(t) !== null
}

/**
 * Excel's Replace in a text: each match of the query, as `matcher` reads it, put as `by` is, the whole text with Match
 * entire cell contents; null where nothing matches or nothing can (an empty query). A `*` takes as little as it can,
 * but the query's last one the rest of the text (`a*b` in `a1b2b` is `a1b`); an empty match puts nothing.
 */
export function replacer(query: string, by: string, o: Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell'>): ((text: string) => string | null) | null {
  const q = o.caseSensitive ? query : folded(query)
  const w = new Wildcard(o.wildcards ? q : q.replace(/[*?~]/g, '~$&'), '')
  if (w.empty) return null
  return (text) => {
    const t = o.caseSensitive ? text : folded(text)
    if (o.wholeCell) return t !== '' && w.test(t) ? by : null
    let out = ''
    let last = 0
    for (let from = 0; from <= t.length; ) {
      const m = w.find(t, from)
      if (!m) break
      // an empty match puts nothing, and the next one is looked for a character on
      if (m[1] === m[0]) {
        from = m[0] + 1
        continue
      }
      out += text.slice(last, m[0]) + by
      last = from = m[1]
    }
    return last ? out + text.slice(last) : null
  }
}
