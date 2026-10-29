/**
 * What match mode 3 takes of PCRE2's syntax in Excel for Mac, a first look without predictions: classes and `\d \w
 * \s \b` over text outside ASCII, anchors and line breaks, the dot over a surrogate pair, quantifiers lazy and
 * possessive, atomic groups, named groups and back references, lookaround, options at the start and inside, `\Q…\E`,
 * recursion, conditional groups, verbs, start-of-pattern options, patterns that are no regular expression, and a
 * pattern that backtracks heavily over texts of growing length. Each case looks in one cell (1, #N/A or #VALUE!).
 * Answers: `test/excel/cases/xmatch-regex-syntax.json`.
 */
import type { SheetCells } from '../probes.js'
import { heldGroups } from './arrays/common.js'

const as = (n: number) => 'a'.repeat(n) + 'b'

/** Rp: the texts the cases look in, line breaks made by CHAR. */
export const SHEETS: SheetCells = {
  Rp: {
    A1: 'abc', A2: 'aaa', A3: 'concat', A4: 'cat food', A5: 'café', A6: '٣', A7: 'a b',
    A8: { formula: '"a"&CHAR(10)&"b"' }, A9: { formula: '"a"&CHAR(13)&CHAR(10)&"b"' }, A10: { formula: '"ab"&CHAR(10)' },
    A11: '\u{1F600}', A12: 'Ä', A13: 'Д', A14: 'K', A15: 'a.b*c', A16: 'abab', A17: '(a(b)c)',
    A18: 'a{,2}', A19: 'x y', A20: as(16), A21: as(20), A22: as(24), A23: { formula: '"a"&CHAR(13)&"b"' },
    A24: 'aa', A25: 'bb', A26: 'ab', A27: 'c', A28: 'C', A29: '42', A30: 'é', A31: 'a{', A32: as(28),
  },
}

/** `[group, pattern, cell of Rp]`. */
const CASES: [string, string, string][] = [
  ['class', '[a-c]+', 'A1'], ['class', '[^a-c]', 'A1'], ['class', '[[:alpha:]]', 'A1'], ['class', '[[:digit:]]', 'A6'],
  ['class', '\\d', 'A6'], ['class', '^caf\\w$', 'A5'], ['class', 'a\\sb', 'A7'], ['class', 'a\\hb', 'A7'],
  ['class', '\\bcat\\b', 'A3'], ['class', '\\bcat\\b', 'A4'], ['class', 'caf\\b', 'A5'], ['class', '\\p{L}', 'A13'],
  ['class', '\\p{Cyrillic}', 'A13'], ['class', '\\P{L}', 'A13'], ['class', '[]a]', 'A1'], ['class', '[\\d\\p{Lu}]', 'A13'],
  ['class', '\\x{414}', 'A13'], ['class', '\\x41', 'A12'], ['class', '\\N{U+414}', 'A13'], ['class', '\\u0414', 'A13'],
  ['class', '[[:^alpha:]]', 'A1'], ['class', '[[:word:]]+', 'A1'],
  ['anchor', '\\Aabc\\z', 'A1'], ['anchor', 'b\\Z', 'A10'], ['anchor', 'b\\z', 'A10'], ['anchor', 'b$', 'A10'],
  ['anchor', '^b', 'A8'], ['anchor', '(?m)^b', 'A8'], ['anchor', 'a$', 'A8'], ['anchor', '(?m)a$', 'A8'],
  ['anchor', '(?m)a$', 'A9'], ['anchor', '\\Gabc', 'A1'], ['anchor', '\\Gb', 'A1'],
  ['dot', 'a.b', 'A8'], ['dot', '(?s)a.b', 'A8'], ['dot', 'a..b', 'A9'], ['dot', 'a.b', 'A23'], ['dot', 'a\\Rb', 'A9'],
  ['dot', 'a\\Nb', 'A8'], ['dot', '^.$', 'A11'], ['dot', '^..$', 'A11'], ['dot', '^\\x{1F600}$', 'A11'],
  ['dot', '^[\u{1F600}]$', 'A11'], ['dot', '^\\X$', 'A30'], ['dot', '^.$', 'A30'],
  ['quant', 'a{2}', 'A2'], ['quant', '^a{,2}$', 'A24'], ['quant', '^a{,2}$', 'A18'], ['quant', 'a+?', 'A2'],
  ['quant', 'a++b', 'A2'], ['quant', '^a*+a', 'A2'], ['quant', '^a*a', 'A2'], ['quant', 'a{', 'A31'],
  ['quant', 'a{2,1}', 'A2'], ['quant', '(?>a+)a', 'A2'], ['quant', 'a**', 'A2'], ['quant', 'a{2}{2}', 'A2'],
  ['quant', '^a?+a', 'A24'], ['quant', '(?U)^a+$', 'A2'],
  ['group', '(a)\\1', 'A2'], ['group', '(?<x>a)\\k<x>', 'A2'], ['group', "(?'x'a)\\k'x'", 'A2'], ['group', '(?P<x>a)(?P=x)', 'A2'],
  ['group', '(a)\\g{-1}', 'A2'], ['group', '(a)\\g1', 'A2'], ['group', '(?<x>a)\\k{x}', 'A2'], ['group', '^(?|(a)|(b))\\1$', 'A25'],
  ['group', '^(ab){2}$', 'A16'], ['group', '(?<x>a)|(?<x>b)', 'A1'], ['group', '(?J)(?<x>a)|(?<x>b)', 'A1'], ['group', '(a)\\2', 'A2'],
  ['group', '(?n)(a)\\1', 'A2'], ['group', '(?#note)abc', 'A1'], ['group', 'a|', 'A27'],
  ['look', 'a(?=b)', 'A1'], ['look', 'a(?!b)', 'A1'], ['look', '(?<=a)b', 'A1'], ['look', '(?<!a)b', 'A1'],
  ['look', '(?<=ab|c)c', 'A1'], ['look', '(?<=a+)b', 'A1'], ['look', '(?<=a(?=b))b', 'A1'],
  ['option', 'a(?i)BC', 'A1'], ['option', 'a(?i:B)c', 'A1'], ['option', '(?i)a(?-i)B', 'A1'], ['option', '(?i)ä', 'A12'],
  ['option', '(?i)д', 'A13'], ['option', '(?i)k', 'A14'], ['option', '^(a(?i)b|c)$', 'A28'], ['option', '(?x)a b c', 'A1'],
  ['option', '(?x)x\\ y', 'A19'], ['option', '(?x)ab#c', 'A1'], ['option', '\\Qa.b*c\\E', 'A15'], ['option', '\\Q.\\E', 'A1'],
  ['option', '(*UCP)\\d', 'A6'], ['option', '(*UTF)^.$', 'A11'], ['option', '(*CR)a.b', 'A23'], ['option', '(*CRLF)(?m)a$', 'A9'],
  ['option', '(?xx)[a b]+c', 'A1'],
  ['recurse', '\\((?:[^()]|(?R))*\\)', 'A17'], ['recurse', '^(\\((?:[^()]|(?1))*\\))$', 'A17'], ['recurse', '^(?<p>\\((?:[^()]|(?&p))*\\))$', 'A17'],
  ['recurse', '^(?P<p>\\((?:[^()]|(?P>p))*\\))$', 'A17'], ['recurse', '^(\\((?:[^()]|\\g<1>)*\\))$', 'A17'],
  ['cond', '^(a)?(?(1)b|c)$', 'A26'], ['cond', '^(a)?(?(1)b|c)$', 'A27'], ['cond', '(?(DEFINE)(?<d>\\d))^(?&d)+$', 'A29'],
  ['cond', '^(?(?=a)ab|c)$', 'A26'], ['cond', '^(?<x>a)?(?(<x>)b|c)$', 'A26'], ['cond', '(?(R)a|b)', 'A26'],
  ['verb', 'a(*FAIL)|b', 'A1'], ['verb', 'a(*ACCEPT)z', 'A1'], ['verb', 'a+(*COMMIT)b', 'A2'], ['verb', 'a(*SKIP)x|b', 'A1'],
  ['verb', 'a(*PRUNE)x|b', 'A1'], ['verb', '(?:a(*THEN)x|ab)', 'A1'], ['verb', '(*MARK:m)abc', 'A1'], ['verb', 'a\\Kb', 'A1'],
  ['verb', '(?C1)abc', 'A1'], ['verb', '\\Cb', 'A1'],
  ['bad', '\\', 'A1'], ['bad', 'a)', 'A1'], ['bad', '*a', 'A1'], ['bad', '(?<=a*)b', 'A1'], ['bad', '[b-a]', 'A1'], ['bad', '(?z)a', 'A1'],
  ['backtrack', '(a+)+$', 'A20'], ['backtrack', '(a+)+$', 'A21'], ['backtrack', '(a+)+$', 'A22'], ['backtrack', '(a+)+$', 'A32'],
]

/** The set's cases (no predictions: a first look). */
export const { buildProbes } = heldGroups(CASES.map(([g, p, cell]) => [g, 'XMATCH', `XMATCH("${p}",Rp!${cell},3)`]))
