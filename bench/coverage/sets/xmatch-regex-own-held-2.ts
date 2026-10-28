/**
 * Match mode 3 with the PCRE2 engine after `xmatch-regex-own-held`'s answers became its rules, predicted by it before
 * Excel for Mac ran the set: the match limit for shapes neither set had (`(a|a)+$`, `(\d+)*$`, a group without
 * capture, a lazy repeat, alternatives of different lengths), one letter under it and at it; the required character in
 * a text of 32,000; the dotted capital I; `(?m)^` and `$` around a newline at the end.
 * Answers: `test/excel/cases/xmatch-regex-own-held-2.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldGroups } from './arrays/common.js'
import { e, n } from '../vals.js'

const a = (k: number) => 'a'.repeat(k) + 'b'

/** Rk: the texts the cases look in. */
export const SHEETS: SheetCells = {
  Rk: {
    A1: a(17), A2: a(18), A3: '1'.repeat(18) + 'x', A4: '1'.repeat(19) + 'x', A5: a(19), A6: 'ab'.repeat(16) + 'c', A7: 'ab'.repeat(17) + 'c',
    A8: { formula: 'REPT("a",32000)&"c"' }, A9: 'İ', A10: 'i', A11: { formula: '"a"&CHAR(13)&CHAR(10)' }, A12: { formula: '"a"&CHAR(13)' },
  },
}

/** `[group, formula, the engine's prediction]`. */
const CASES: [string, string, Val][] = [
  ['limit', 'XMATCH("(a|a)+$",Rk!A1,3)', e('NA')],
  ['limit', 'XMATCH("(a|a)+$",Rk!A2,3)', e('VALUE')],
  ['limit', 'XMATCH("(\\d+)*$",Rk!A3,3)', n(1)],
  ['limit', 'XMATCH("(\\d+)*$",Rk!A4,3)', e('VALUE')],
  ['limit', 'XMATCH("(?:a+)+$",Rk!A2,3)', e('NA')],
  ['limit', 'XMATCH("(?:a+)+$",Rk!A5,3)', e('VALUE')],
  ['limit', 'XMATCH("(a+?)+$",Rk!A2,3)', e('NA')],
  ['limit', 'XMATCH("(a+?)+$",Rk!A5,3)', e('VALUE')],
  ['limit', 'XMATCH("(a|b|ab)+$",Rk!A6,3)', e('NA')],
  ['limit', 'XMATCH("(a|b|ab)+$",Rk!A7,3)', e('VALUE')],
  ['required', 'XMATCH("(a+)+b",Rk!A8,3)', e('NA')],
  ['case', 'XMATCH("(?i)İ",Rk!A10,3)', e('NA')],
  ['case', 'XMATCH("(?i)i",Rk!A9,3)', e('NA')],
  ['line', 'XMATCH("(?m)^$",Rk!A11,3)', n(1)],
  ['line', 'XMATCH("(?m)a$",Rk!A12,3)', n(1)],
  ['line', 'XMATCH("(?m)^$",Rk!A12,3)', n(1)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldGroups(CASES.map(([g, f, v]) => [g, 'XMATCH', f, v]))
