/**
 * TEXTAFTER, TEXTBEFORE and TEXTSPLIT on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by
 * Excel) that the engine answered otherwise: a fractional instance from the end (-1.6, -1.2), a delimiter longer than
 * the text with if_not_found given, TEXTSPLIT with neither delimiter. Then the rules around them: fractional instances
 * both ways and a hair under a whole number, delimiters as long as the text, one of several longer, case-insensitive,
 * match_end; TEXTSPLIT with a row delimiter alone. Answers: `test/excel/cases/textafter-ironcalc-held.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

const T = '"123 the 456 the 891"'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['TEXTAFTER', `TEXTAFTER(${T},"the",-1.6,1,0,"Not found")`, s(' 456 the 891')],
  ['TEXTAFTER', `TEXTAFTER(${T},"the",-1.2,1,0,"Not found")`, s(' 456 the 891')],
  ['TEXTBEFORE', `TEXTBEFORE(${T},"the",-1.6,1,0, "Not found")`, s('123 ')],
  ['TEXTAFTER', 'TEXTAFTER("del","deli",1,1,0,"qwerty")', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("hi","Hello Good Morning",1,1,0,"@")', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("del","deli",1)', e('VALUE')],
  ['TEXTBEFORE', 'TEXTBEFORE("del","deli",1,1,0, "qwerty")', e('VALUE')],
  ['TEXTSPLIT', 'TEXTSPLIT("Hola sin split",)', e('VALUE')],
  ['TEXTSPLIT', 'TEXTSPLIT("Hola sin split",,)', e('VALUE')],
  // fractional instances, rounded down, and a hair under a whole number
  ['TEXTAFTER', `TEXTAFTER(${T},"the",1.6)`, s(' 456 the 891')],
  ['TEXTAFTER', `TEXTAFTER(${T},"the",-0.5)`, s(' 891')],
  ['TEXTAFTER', `TEXTAFTER(${T},"the",0.5)`, e('VALUE')],
  ['TEXTAFTER', `TEXTAFTER(${T},"the",1.9999999)`, s(' 891')],
  ['TEXTBEFORE', `TEXTBEFORE(${T},"the",-2.5)`, e('NA')],
  ['TEXTBEFORE', `TEXTBEFORE(${T},"the",2.7)`, s('123 the 456 ')],
  // delimiters as long as the text, one of several longer, in any case, with match_end
  ['TEXTAFTER', 'TEXTAFTER("abc","abc")', s('')],
  ['TEXTAFTER', 'TEXTAFTER("abc",{"abcd","b"})', s('c')],
  ['TEXTAFTER', 'TEXTAFTER("abc",{"abcd","xyzw"},1,0,0,"-")', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("abc","ABCD",1,1)', e('VALUE')],
  ['TEXTBEFORE', 'TEXTBEFORE("abc","abcd",1,0,1)', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("abc","x",1,0,0,"-")', s('-')],
  // TEXTSPLIT with a row delimiter alone, and a column one
  ['TEXTSPLIT', 'ROWS(TEXTSPLIT("a;b",,";"))', n(2)],
  ['TEXTSPLIT', 'COLUMNS(TEXTSPLIT("a b"," "))', n(2)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * All as predicted, but the engine's rule broke a recorded case: `TEXTBEFORE("A","aBc")` is #N/A in Excel, so a
 * delimiter longer than the text is #VALUE! only once an instance is given. The set `textafter-ironcalc-held-2`
 * measures that.
 */
export const MISSED: readonly string[] = []
