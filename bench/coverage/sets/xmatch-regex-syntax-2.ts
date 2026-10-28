/**
 * Match mode 3, what `xmatch-regex-syntax` left open, again without predictions: which characters end a line (NEL,
 * the line separator, VT, FF), a lookbehind of bounded length, `\w \d \s` over a combining mark, a connector, a
 * fullwidth digit and the line separator, caseless Greek sigma and a back reference, options of newer PCRE2
 * (`(?^)`, `(?aD)`, `(*LIMIT_MATCH=…)`), relative recursion, a lone surrogate, and where the match limit lies: the same
 * heavy pattern over texts of 17 to 19 letters, another over growing texts, and many short runs in one text.
 * Answers: `test/excel/cases/xmatch-regex-syntax-2.json`.
 */
import type { SheetCells } from '../probes.js'
import { heldGroups } from './arrays/common.js'

const as = (n: number) => 'a'.repeat(n) + 'b'
const between = (code: number) => ({ formula: `"a"&_xlfn.UNICHAR(${code})&"b"` })

/** Rq: the texts the cases look in. */
export const SHEETS: SheetCells = {
  Rq: {
    A1: between(133), A2: between(8232), A3: between(11), A4: between(12), A5: 'abc', A6: 'é', A7: 'a‿b',
    A8: '３', A9: 'ς', A10: 'aA', A11: { formula: '"ab"&CHAR(13)&CHAR(10)' }, A12: 'a', A13: '٣',
    A14: as(17), A15: as(18), A16: as(19), A17: as(23), A18: as(27), A19: as(31), A20: as(16).repeat(40), A21: '(a(b)c)',
    A22: 'ab', A23: { formula: '"a"&CHAR(13)' }, A24: 'a_b', A25: as(24), A26: as(25), A27: as(26),
    A28: 'x'.repeat(10) + 'zy', A29: 'x'.repeat(12) + 'zy', A30: 'x'.repeat(14) + 'zy', A31: 'x'.repeat(16) + 'zy',
  },
}

/** `[group, pattern, cell of Rq]`. */
const CASES: [string, string, string][] = [
  ['line', 'a.b', 'A1'], ['line', '(?m)a$', 'A1'], ['line', 'a\\Rb', 'A1'], ['line', 'a.b', 'A2'], ['line', '(?m)a$', 'A2'],
  ['line', 'a.b', 'A3'], ['line', 'a.b', 'A4'], ['line', 'b$', 'A11'], ['line', 'a$', 'A23'], ['line', 'a\\vb', 'A1'],
  ['line', '(*LF)a.b', 'A1'], ['line', '(*ANYCRLF)a.b', 'A1'],
  ['look', '(?<=a{1,3})b', 'A5'], ['look', '(?<=a|bc)c', 'A5'], ['look', '(?<=(a)|b)c', 'A5'],
  ['class', '^e\\w$', 'A6'], ['class', '^a\\wb$', 'A7'], ['class', '\\d', 'A8'], ['class', 'a\\sb', 'A2'], ['class', '^\\w+$', 'A24'],
  ['class', '[[:punct:]]', 'A7'], ['class', '\\p{Xan}', 'A8'], ['class', '\\p{sc=Cyrl}', 'A5'], ['class', '\\p{L&}', 'A9'],
  ['class', '\\p{Greek}', 'A9'], ['class', '\\p{ greek }', 'A9'], ['class', '\\pL', 'A9'], ['class', '\\p{Bidi_Class:L}', 'A9'],
  ['case', '(?i)σ', 'A9'], ['case', '(?i)(a)\\1', 'A10'], ['case', '(?i)[^a]', 'A10'], ['case', '(?i)\\p{Lu}', 'A12'],
  ['case', '(?i)[A-Z]', 'A12'],
  ['option', '(?i)(?^)A', 'A12'], ['option', '(?aD)\\d', 'A13'], ['option', '(*LIMIT_MATCH=100)(a+)+$', 'A14'],
  ['option', '(*NO_AUTO_POSSESS)a+b', 'A5'], ['option', '(*UTF)(*UCP)abc', 'A5'], ['option', '(?-x)a', 'A5'],
  ['recurse', '^(\\((?:[^()]|(?-1))*\\))$', 'A21'], ['recurse', '^(?+1)(\\((?:[^()]|(?1))*\\))$', 'A21'], ['recurse', '^(a|b(?1))$', 'A22'],
  ['bad', '\\x{D800}', 'A5'], ['bad', '(?<=a(*ACCEPT))', 'A5'], ['bad', '\\p{Nope}', 'A5'], ['bad', 'a{65536}', 'A5'], ['bad', '(?', 'A5'],
  ['backtrack', '(a+)+$', 'A14'], ['backtrack', '(a+)+$', 'A15'], ['backtrack', '(a+)+$', 'A16'], ['backtrack', '(a|aa)+$', 'A17'],
  ['backtrack', '(a|aa)+$', 'A18'], ['backtrack', '(a|aa)+$', 'A19'], ['backtrack', '(a+)+$', 'A20'], ['backtrack', '^(a+)+$', 'A16'],
  ['backtrack', '(a|aa)+$', 'A25'], ['backtrack', '(a|aa)+$', 'A26'], ['backtrack', '(a|aa)+$', 'A27'], ['backtrack', '(x+x+)+y', 'A28'],
  ['backtrack', '(x+x+)+y', 'A29'], ['backtrack', '(x+x+)+y', 'A30'], ['backtrack', '(x+x+)+y', 'A31'],
  ...[200_000, 500_000, 1_000_000, 2_000_000, 5_000_000, 10_000_000].map((n): [string, string, string] => ['limit', `(*LIMIT_MATCH=${n})(a+)+$`, 'A15']),
]

/** The set's cases (no predictions: a first look). */
export const { buildProbes } = heldGroups(CASES.map(([g, p, cell]) => [g, 'XMATCH', `XMATCH("${p}",Rq!${cell},3)`]))
