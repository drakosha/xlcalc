/**
 * ENCODEURL held out, a set of Excel for Windows alone (Excel for Mac 16.114 has no ENCODEURL): which characters it
 * keeps, UTF-8 bytes of accented letters, CJK and characters past U+FFFF, numbers and booleans as text, a lone
 * surrogate, an answer longer than 32,767 characters. \`PREDICTED\` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

/** W: a text with a space and a number. */
export const SHEETS: SheetCells = { W: { A1: 'x y', A2: 42 } }

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // letters, digits and -_.~ kept, every other character as its UTF-8 bytes
  ['ENCODEURL', 'ENCODEURL("a b")', s("a%20b")],
  ['ENCODEURL', "ENCODEURL(\"!*'();:@&=+$,/?#[]\")", s("%21%2A%27%28%29%3B%3A%40%26%3D%2B%24%2C%2F%3F%23%5B%5D")],
  ['ENCODEURL', 'ENCODEURL("~-_.")', s("~-_.")],
  ['ENCODEURL', 'ENCODEURL("%")', s("%25")],
  ['ENCODEURL', 'ENCODEURL("é")', s("%C3%A9")],
  ['ENCODEURL', 'ENCODEURL("中文")', s("%E4%B8%AD%E6%96%87")],
  ['ENCODEURL', 'ENCODEURL("Ünïcödé")', s("%C3%9Cn%C3%AFc%C3%B6d%C3%A9")],
  ['ENCODEURL', 'ENCODEURL("ÿ")', s("%C3%BF")],
  ['ENCODEURL', 'ENCODEURL("a""b")', s("a%22b")],
  ['ENCODEURL', 'ENCODEURL("")', s("")],
  ['ENCODEURL', 'ENCODEURL("ABCxyz019")', s("ABCxyz019")],
  // values of other kinds, a lone surrogate, cells, the length of the answer, an array
  ['ENCODEURL', 'ENCODEURL(1.5)', s("1.5")],
  ['ENCODEURL', 'ENCODEURL(TRUE)', s("TRUE")],
  ['ENCODEURL', 'ENCODEURL(1E+21)', s("1E%2B21")],
  ['ENCODEURL', 'ENCODEURL(NA())', e('NA')],
  ['ENCODEURL', 'ENCODEURL(UNICHAR(128512))', s("%F0%9F%98%80")],
  ['ENCODEURL', 'ENCODEURL(CHAR(10)&CHAR(9))', s("%0A%09")],
  ['ENCODEURL', 'ENCODEURL(MID(UNICHAR(128512),1,1))', e('VALUE')],
  ['ENCODEURL', 'ENCODEURL(MID(UNICHAR(128512),2,1))', e('VALUE')],
  ['ENCODEURL', 'ENCODEURL(W!A1)', s("x%20y")],
  ['ENCODEURL', 'ENCODEURL(W!A2)', s("42")],
  ['ENCODEURL', 'LEN(ENCODEURL(REPT("a",2048)))', n(2048)],
  ['ENCODEURL', 'LEN(ENCODEURL(REPT("a",2049)))', n(2049)],
  ['ENCODEURL', 'LEN(ENCODEURL(REPT("é",5461)))', n(32766)],
  ['ENCODEURL', 'LEN(ENCODEURL(REPT("é",5462)))', e('VALUE')],
  ['ENCODEURL', 'LEN(ENCODEURL(REPT(" ",10923)))', e('VALUE')],
  ['ENCODEURL', 'ENCODEURL({"a b","c"})', e('SPILL')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
