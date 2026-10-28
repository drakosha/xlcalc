/**
 * Match mode 3 with the PCRE2 engine of `xlcalc/regex` (`pcreEngine`), predicted by it before Excel for Mac ran the
 * set: where Excel's match limit falls for heavy patterns of other shapes than the calibration's (`(x+x+)+y`,
 * `(a*)*b`, words and spaces, `(a|a?)+$`, a caseless and a class repeat), the required character a match must hold
 * (missing: #N/A at once, in texts of up to 5,000 characters), the limit inside a range and in `XLOOKUP`, caseless
 * letters outside ASCII, `\X`, `\b` in Cyrillic, a lookbehind of 255 and 256 characters, the verbs that cut
 * backtracking, a recursive palindrome, a back reference to an unset group, duplicate names, `(*NOTEMPTY)`.
 * Answers: `test/excel/cases/xmatch-regex-own-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldGroups } from './arrays/common.js'
import { e, n } from '../vals.js'

const a = (k: number, end = 'b') => 'a'.repeat(k) + end
const words = (k: number) => Array(k).fill('ab').join(' ') + '!'
const x = (k: number) => 'x'.repeat(k) + 'zy'

/** Rh: the texts the cases look in. */
export const SHEETS: SheetCells = {
  Rh: {
    A1: x(19), A2: x(20), A3: x(21), A4: a(18, 'cb'), A5: a(19, 'cb'), A6: a(20, 'cb'), A7: a(30, ''), A8: words(16), A9: words(17), A10: words(18),
    A11: a(16), A12: a(17), A13: a(18), A14: a(30, 'c'), A15: a(19), A16: a(18, '1'), A17: a(19, '1'), A18: 'aa',
    A19: { formula: 'REPT("a",1500)&"c"' }, A20: { formula: 'REPT("a",4000)&"c"' }, A21: { formula: 'REPT("a",6000)&"c"' },
    A22: 'ǆ', A23: 'ẞ', A24: 'I', A25: 'i', A26: '\u{1f468}‍\u{1f469}‍\u{1f467}', A27: 'привет мир', A28: 'ac',
    A29: 'abcba', A30: 'b', A31: 'bb', A32: { formula: '"a"&CHAR(10)' }, A33: 'ba', A34: 'a'.repeat(255) + 'b',
    B1: 1, B2: 2,
  },
}

/** `[group, formula, the engine's prediction]`. */
const CASES: [string, string, Val][] = [
  // the match limit by shape, around the engine's threshold
  ['limit', 'XMATCH("(x+x+)+y",Rh!A1,3)', e('NA')],
  ['limit', 'XMATCH("(x+x+)+y",Rh!A2,3)', e('VALUE')],
  ['limit', 'XMATCH("(x+x+)+y",Rh!A3,3)', e('VALUE')],
  ['limit', 'XMATCH("(a*)*b",Rh!A4,3)', n(1)],
  ['limit', 'XMATCH("(a*)*b",Rh!A5,3)', e('VALUE')],
  ['limit', 'XMATCH("(a*)*b",Rh!A6,3)', e('VALUE')],
  ['limit', 'XMATCH("^(\\w+\\s?)+$",Rh!A8,3)', e('NA')],
  ['limit', 'XMATCH("^(\\w+\\s?)+$",Rh!A9,3)', e('VALUE')],
  ['limit', 'XMATCH("^(\\w+\\s?)+$",Rh!A10,3)', e('VALUE')],
  ['limit', 'XMATCH("(a|a?)+$",Rh!A11,3)', n(1)],
  ['limit', 'XMATCH("(a|a?)+$",Rh!A12,3)', e('VALUE')],
  ['limit', 'XMATCH("(a|a?)+$",Rh!A13,3)', e('VALUE')],
  ['limit', 'XMATCH("(?i)(A+)+$",Rh!A13,3)', e('NA')],
  ['limit', 'XMATCH("(?i)(A+)+$",Rh!A15,3)', e('VALUE')],
  ['limit', 'XMATCH("([a-z]+)+$",Rh!A16,3)', e('NA')],
  ['limit', 'XMATCH("([a-z]+)+$",Rh!A17,3)', e('VALUE')],
  ['limit', 'XMATCH("(a+)+b",Rh!A4,3)', e('NA')],
  ['limit', 'XMATCH("(a+)+b",Rh!A5,3)', e('VALUE')],
  // a required character missing: no match, without trying
  ['required', 'XMATCH("(a*)*b",Rh!A7,3)', e('NA')],
  ['required', 'XMATCH("(a+)+b",Rh!A14,3)', e('NA')],
  ['required', 'XMATCH("(a+)+b",Rh!A19,3)', e('NA')],
  ['required', 'XMATCH("(a+)+b",Rh!A20,3)', e('NA')],
  ['required', 'XMATCH("(a+)+b",Rh!A21,3)', e('VALUE')],
  // the limit in a range and in XLOOKUP
  ['range', 'XMATCH("(a+)+$",Rh!A15:A18,3)', e('VALUE')],
  ['range', 'XMATCH("(a+)+$",Rh!A18:A15,3)', e('VALUE')],
  ['range', 'XMATCH("(a+)+$",Rh!A15:A18,3,-1)', n(4)],
  ['range', 'XLOOKUP("(a+)+$",Rh!A15:A16,Rh!B1:B2,"none",3)', e('VALUE')],
  // caseless letters outside ASCII
  ['case', 'XMATCH("(?i)ǅ",Rh!A22,3)', n(1)],
  ['case', 'XMATCH("(?i)ß",Rh!A23,3)', n(1)],
  ['case', 'XMATCH("(?i)ı",Rh!A24,3)', n(1)],
  ['case', 'XMATCH("(?i)ı",Rh!A25,3)', n(1)],
  // \X, \b in Cyrillic, lookbehinds of 255 and 256 characters
  ['class', 'XMATCH("^\\X$",Rh!A26,3)', n(1)],
  ['class', 'XMATCH("\\bмир\\b",Rh!A27,3)', n(1)],
  ['class', 'XMATCH("\\Bир\\b",Rh!A27,3)', n(1)],
  ['look', 'XMATCH("(?<=a{1,255})b",Rh!A34,3)', n(1)],
  ['look', 'XMATCH("(?<=a{1,256})b",Rh!A34,3)', e('VALUE')],
  ['look', 'XMATCH("(?<=a{255})b",Rh!A34,3)', n(1)],
  // the verbs that cut backtracking
  ['verb', 'XMATCH("a(*COMMIT)b|ac",Rh!A28,3)', e('NA')],
  ['verb', 'XMATCH("a(*PRUNE)b|ac",Rh!A28,3)', e('NA')],
  ['verb', 'XMATCH("a(*THEN)b|ac",Rh!A28,3)', n(1)],
  ['verb', 'XMATCH("a(*SKIP)b|c",Rh!A28,3)', n(1)],
  // recursion, an unset group, duplicate names, empty matches
  ['group', 'XMATCH("^((.)(?1)\\2|.?)$",Rh!A29,3)', n(1)],
  ['group', 'XMATCH("(a)?b\\1",Rh!A30,3)', e('NA')],
  ['group', 'XMATCH("(?J)(?:(?<n>a)|(?<n>b))\\k<n>",Rh!A31,3)', n(1)],
  ['group', 'XMATCH("(?m)^$",Rh!A32,3)', e('NA')],
  ['group', 'XMATCH("^(a|)*$",Rh!A18,3)', n(1)],
  ['group', 'XMATCH("(*NOTEMPTY)a*",Rh!A30,3)', e('NA')],
  ['group', 'XMATCH("(*NOTEMPTY_ATSTART)a*",Rh!A33,3)', n(1)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldGroups(CASES.map(([g, f, v]) => [g, 'XMATCH', f, v]))
