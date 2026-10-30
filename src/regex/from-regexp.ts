/** A JavaScript regular expression as a level's pattern (`RegexPattern`): the built-in level's, `pcreSyntax`'s. */
import type { RegexPattern } from '../contract.js'

/**
 * `re` as a pattern; `groups` gives the spans of the pattern's own groups from a match's indices (its numbered ones
 * for a pattern in JavaScript's syntax, those `pcreSyntax` names after them).
 */
export function regExpPattern(re: RegExp, groups: (d: RegExpIndicesArray) => readonly ([number, number] | undefined)[]): RegexPattern {
  let all: RegExp | undefined
  return {
    test: (s) => re.test(s),
    exec(s, from) {
      all ??= new RegExp(re.source, `${re.flags}gd`)
      all.lastIndex = from
      const m = all.exec(s)
      if (!m) return null
      const out = [m.index, m.index + m[0].length]
      for (const g of groups(m.indices!)) out.push(g ? g[0] : -1, g ? g[1] : -1)
      return out
    },
  }
}
