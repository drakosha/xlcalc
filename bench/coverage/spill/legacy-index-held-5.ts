/**
 * Held out after `legacy-index-held-4` (`run.ts spill legacy-index-held-5`): forty formulas of the old kind that tell
 * the rules of INDEX's place apart (MATCH against the other position functions, MATCH under an operator or a call,
 * MATCH with no match, a range against an array, names of a range, a formula and constants, literals at 65535 and
 * 65536, a one-cell source), and four literals spelled otherwise. `SHOWN` is the text predicted before Excel ran, by
 * the rule that only a whole number literal up to 65535 (in parentheses too), MATCH itself and a name of such a
 * literal leave one cell; `CONVERTED` is the converter's text at the same commit.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { fromLegacy } from '../../../src/xlsx.js'
import { canonical, parse } from '../../../src/syntax.js'

const DATA: Record<string, CellSpec> = {
  ...Object.fromEntries([1, 2, 3, 4, 5].map((v, i) => [`A${i + 1}`, v])),
  ...Object.fromEntries([10, 20, 30, 40, 50].map((v, i) => [`B${i + 1}`, v])),
  ...Object.fromEntries([20, 120, 130, 140, 150].map((v, i) => [`C${i + 1}`, v])),
  D1: 2,
}

const NAMES = [
  { name: 'nData', ref: "'S'!$B$1:$B$5" },
  { name: 'nPos', ref: "MATCH(3,'S'!$A$1:$A$5,0)" },
  { name: 'nTwo', ref: '2' },
  { name: 'nZero', ref: '0' },
]

/** Formula and the text predicted; `=` alone means the formula as it is. */
const CASES: [string, string][] = [
  ['INDEX(B1:B5,2)', '='],
  ['INDEX(B1:B5,0)', '@'],
  ['INDEX(B1:B5,-1)', '@'],
  ['INDEX(B1:B5,2.5)', '@'],
  ['INDEX(B1:B5,TRUE)', '@'],
  ['INDEX(B1:B5,"2")', '@'],
  ['INDEX(B1:B5,1+1)', '@'],
  ['INDEX(B1:B5,D1)', '@'],
  ['INDEX(B1:B5,(2))', '='],
  ['INDEX(B1:B5,MATCH(3,A1:A5,0))', '='],
  ['INDEX(B1:B5,MATCH(99,A1:A5,0))', '='],
  ['INDEX(B1:B5,XMATCH(3,A1:A5,0))', '@'],
  ['INDEX(B1:B5,MATCH(3,A1:A5,0)+0)', '@'],
  ['INDEX(B1:B5,ABS(MATCH(3,A1:A5,0)))', '@'],
  ['INDEX(B1:B5,IF(TRUE,MATCH(3,A1:A5,0),1))', '@'],
  ['INDEX(B1:B5,ROW(A2))', '@'],
  ['INDEX(B1:B5,COLUMN(B1))', '@'],
  ['INDEX(B1:B5,ROWS(A1:A3))', '@'],
  ['INDEX(B1:B5,COLUMNS(A1:C1))', '@'],
  ['INDEX(B1:B5,COUNT(A1:A3))', '@'],
  ['INDEX(B1:B5,COUNT(F1:F3))', '@'],
  ['INDEX(B1:B5,SUM(0))', '@'],
  ['INDEX(B1:B5,SUM(2))', '@'],
  ['INDEX(B1:B5,LEN(""))', '@'],
  ['INDEX(B1:B5,INT(2.5))', '@'],
  ['INDEX(B1:B5,FIND("a","abc"))', '@'],
  ['INDEX(B1:B5,RANK(3,A1:A5,0))', '@'],
  ['INDEX(B1:B5,FALSE())', '@'],
  ['INDEX(B1:C5,MATCH(3,A1:A5,0),2)', '='],
  ['INDEX(B1:C5,2,MATCH(20,B1:C1,0))', '='],
  ['INDEX(B1:D1,MATCH(20,B1:D1,0))', '='],
  ['INDEX({10;20;30},MATCH(3,A1:A5,0))', '@'],
  ['INDEX({10;20;30},MATCH(3,A1:A5,0),1)', '='],
  ['INDEX(nData,MATCH(3,A1:A5,0))', '='],
  ['INDEX(B1:B5,nPos)', '@INDEX(B1:B5,@nPos)'],
  ['INDEX(B1:B5,nTwo)', '='],
  ['INDEX(B1:B5,nZero)', '@'],
  ['INDEX(B:B,65535)', '='],
  ['INDEX(B:B,65536)', '@'],
  ['INDEX(B1,0)', '='],
  ['INDEX(B1:B5,2.0)', '='],
  ['INDEX(B1:B5,2E0)', '='],
  ['INDEX(B1:B5,+2)', '@'],
  ['INDEX(B:B,70000)', '@'],
]

/** The cell of case `i` (from 0): row 3 from H on. */
const cellOf = (i: number) => `${String.fromCharCode(...(i + 7 < 26 ? [65 + i + 7] : [64 + Math.floor((i + 7) / 26), 65 + ((i + 7) % 26)]))}3`

const FORMULAS: Record<string, string> = Object.fromEntries(CASES.map(([f], i) => [cellOf(i), `=${f}`]))
const at = Object.keys(FORMULAS)

export function buildSpillCases(): SpillCase[] {
  return [
    one('forty', 'forty formulas that tell the rules of INDEX\'s place apart, of the old kind', { ...DATA, ...FORMULAS }, at, {
      book: { sheets: [], names: NAMES },
      legacy: at.map((a) => `S!${a}`),
    }),
  ]
}

/** The text Excel 365 was predicted to show (canonical, without `=`), committed before Excel ran. */
export const SHOWN: Readonly<Record<string, string>> = Object.fromEntries(CASES.map(([f, t], i) => [`S!${cellOf(i)}`, canonical(parse(t === '=' ? f : t === '@' ? `@${f}` : t))]))

/** The predictions Excel did not confirm: Excel reads `+2` as the literal 2, and shows it so. */
export const MISSED: readonly string[] = ['S!AX3']

/** The converter's text of each formula. */
export const CONVERTED = (): Record<string, string> => Object.fromEntries(at.map((a) => [`S!${a}`, fromLegacy(FORMULAS[a].slice(1), { names: NAMES, sheet: 'S' })]))
