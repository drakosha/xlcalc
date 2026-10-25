/**
 * Excel's wildcard patterns — `*` any text, `?` one character, `~` the next character as it is — as matching: the
 * regular expression a pattern is, a pattern matched whole and by its parts between `*`, and a pattern's literal
 * text escaped. Criteria read them (`src/criteria.ts`), so do lookups and SEARCH (`wildMatcher`), and the `./text`
 * subpath's lexer and Find take the class without the engine: this module imports nothing and computes nothing else.
 */

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
