/**
 * XLOOKUP on the cases of IronCalc's tests (`xlsx/tests/calc_tests/LOOKUP_AND_REFERENCE/XLOOKUP.xlsx`, saved by
 * Excel) that the engine answered otherwise: binary search (search modes 2 and −2) over runs of equal values and
 * over another type than the key's, match modes −1 and 1 across types in a linear search, wildcards with binary
 * search. Then the rule around them, in XLOOKUP and XMATCH: binary search over unsorted data, a blank inside the
 * data, mixed types sorted by type, a whole column searched from the end; the linear next value past blanks and
 * errors. Answers: `test/excel/cases/xlookup-ironcalc-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

const col = (c: string, from: number, values: (number | string | boolean)[]) => Object.fromEntries(values.map((v, i) => [`${c}${from + i}`, v]))
const row = (r: number, from: string, values: (number | string | boolean)[]) =>
  Object.fromEntries(values.map((v, i) => [`${String.fromCharCode(from.charCodeAt(0) + i)}${r}`, v]))
const letters = 'aa a b c d e f g h i j k l'.split(' ')

/** Xl, Me, Rl and Wc: data of the file's sheets LargerRows, MultipleElements, ReturnLarger and Wildcards; the others the set's own. */
export const SHEETS: SheetCells = {
  Xl: { ...row(1, 'B', [2, 3, 4, 5, 6]), ...row(2, 'B', ['a', 'b', 'c', 'd', 'e']), ...row(1, 'Q', [6, 5, 4, 3, 2]), ...row(2, 'Q', ['e', 'd', 'c', 'b', 'a']) },
  Me: { ...col('A', 1, [1, 2, 3, 4, 5, 5, 5, 6, 7, 7, 7, 8, 9]), ...col('B', 1, letters), ...col('I', 1, [9, 8, 7, 7, 7, 6, 5, 5, 5, 4, 3, 2, 1]), ...col('J', 1, letters) },
  Rl: {
    ...col('A', 4, [2, 5, 4, 5, 3, true, 'ala']), ...col('B', 4, ['a', 'b', 'c', 'd', 'e', 'f', 'g']),
    ...col('A', 19, [2, 3, 4, 5, 6]), ...col('B', 19, ['a', 'b', 'c', 'd', 'e']),
    ...col('A', 34, [6, 5, 4, 3, 2]), ...col('B', 34, ['e', 'd', 'c', 'b', 'a']),
  },
  Wc: { ...col('A', 13, ['Abra', 'Bra', 'Cada', 'Fun', 'Dan']), ...col('B', 13, [1, 2, 3, 4, 5]), A20: 'Dan', A21: 'Fun', B20: 1, B21: 2 },
  Us: { ...col('A', 1, [1, 9, 5, 2, 3]), ...col('B', 1, ['u1', 'u2', 'u3', 'u4', 'u5']) },
  Bl: { ...col('A', 1, [1, 2]), ...col('A', 4, [4, 5, 6]), ...col('B', 1, ['b1', 'b2', 'b3', 'b4', 'b5', 'b6']) },
  Mx: { ...col('A', 1, [1, 2, 3, 'a', 'b', true]), ...col('B', 1, ['m1', 'm2', 'm3', 'm4', 'm5', 'm6']) },
  Er: { A1: 1, A2: { formula: 'NA()' }, A4: 3, ...col('B', 1, ['e1', 'e2', 'e3', 'e4']) },
  Dc: { ...col('A', 1, [9, 8, 7, 6, 5]), ...col('B', 1, ['d1', 'd2', 'd3', 'd4', 'd5']) },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases: binary search over text with a number, next larger
  ['XLOOKUP', 'XLOOKUP(12,Xl!$B$2:$F$2,Xl!$B$1:$F$1,,1,2)', n(2)],
  ['XLOOKUP', 'XLOOKUP(123,Xl!$Q$2:$U$2,Xl!$Q$1:$U$1,,1,-2)', n(2)],
  ['XLOOKUP', 'XLOOKUP(FALSE,Xl!$Q$2:$U$2,Xl!$Q$1:$U$1,,-1,-2)', n(6)],
  // runs of equal values: the exact match is the first of the run (the last from the end), the next smaller the last (the first)
  ['XLOOKUP', 'XLOOKUP(7.1,Me!A:A,Me!B:B,,-1,2)', s('j')],
  ['XLOOKUP', 'XLOOKUP(7.1,Me!$I$1:$I$13,Me!$J$1:$J$13,,-1,-2)', s('b')],
  ['XLOOKUP', 'XLOOKUP(7,Me!A:A,Me!B:B,,-1,2)', s('h')],
  ['XLOOKUP', 'XLOOKUP(7,Me!$I$1:$I$13,Me!$J$1:$J$13,,-1,-2)', s('d')],
  // linear search, next larger and smaller across types
  ['XLOOKUP', 'XLOOKUP(10,Rl!$A$4:$A$10,Rl!$B$4:$B$10,,1)', s('g')],
  ['XLOOKUP', 'XLOOKUP("qwerty",Rl!$A$4:$A$10,Rl!$B$4:$B$10,,1,-1)', s('f')],
  ['XLOOKUP', 'XLOOKUP(FALSE,Rl!$A$4:$A$10,Rl!$B$4:$B$10, "Not found", -1)', s('g')],
  // binary search over numbers with text and a boolean
  ['XLOOKUP', 'XLOOKUP(12,Rl!$B$34:$B$38,Rl!$A$34:$A$38,"-",1,-2)', n(2)],
  ['XLOOKUP', 'XLOOKUP("a",Rl!$A$19:$A$23,Rl!$B$19:$B$23,,-1,2)', s('e')],
  ['XLOOKUP', 'XLOOKUP(TRUE,Rl!$B$34:$B$38,Rl!$A$34:$A$38,,-1,-2)', n(6)],
  // wildcards and binary search
  ['XLOOKUP', 'XLOOKUP("Abra",Wc!A13:A20,Wc!B13:B20,,2,2)', e('VALUE')],
  ['XLOOKUP', 'XLOOKUP("Dan",Wc!A20:A21,Wc!B20:B21,,2,-2)', e('VALUE')],
  // the same in XMATCH
  ['XMATCH', 'XMATCH("Abra",Wc!A13:A17,2,2)', e('VALUE')],
  ['XMATCH', 'XMATCH(7.1,Me!A:A,-1,2)', n(11)],
  ['XMATCH', 'XMATCH(7,Me!I1:I13,0,-2)', n(5)],
  ['XMATCH', 'XMATCH(12,Xl!B2:F2,1,2)', n(1)],
  ['XMATCH', 'XMATCH("b",Mx!A1:A6,0,2)', n(5)],
  // binary search over unsorted data (1, 9, 5, 2, 3) and over a blank inside the data (1, 2, blank, 4, 5, 6)
  ['XLOOKUP', 'XLOOKUP(5,Us!A1:A5,Us!B1:B5,,0,2)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(4,Us!A1:A5,Us!B1:B5,,-1,2)', s('u1')],
  ['XLOOKUP', 'XLOOKUP(2,Us!A1:A5,Us!B1:B5,,1,2)', s('u2')],
  ['XLOOKUP', 'XLOOKUP(5,Bl!A1:A6,Bl!B1:B6,,0,2)', s('b5')],
  ['XLOOKUP', 'XLOOKUP(2.5,Bl!A1:A6,Bl!B1:B6,,-1,2)', s('b2')],
  ['XLOOKUP', 'XLOOKUP(2.5,Bl!A1:A6,Bl!B1:B6,,1,2)', s('b3')],
  // mixed types sorted by type (1, 2, 3, "a", "b", TRUE): binary and linear
  ['XLOOKUP', 'XLOOKUP("b",Mx!A1:A6,Mx!B1:B6,,0,2)', s('m5')],
  ['XLOOKUP', 'XLOOKUP(10,Mx!A1:A6,Mx!B1:B6,,1,2)', s('m4')],
  ['XLOOKUP', 'XLOOKUP("z",Mx!A1:A6,Mx!B1:B6,,1,2)', s('m6')],
  ['XLOOKUP', 'XLOOKUP(FALSE,Mx!A1:A6,Mx!B1:B6,,-1,2)', s('m5')],
  ['XLOOKUP', 'XLOOKUP(10,Mx!A1:A6,Mx!B1:B6,,1)', s('m4')],
  ['XLOOKUP', 'XLOOKUP("zz",Mx!A1:A6,Mx!B1:B6,,-1)', s('m5')],
  ['XLOOKUP', 'XLOOKUP("zz",Mx!A1:A6,Mx!B1:B6,,1)', s('m6')],
  ['XLOOKUP', 'XLOOKUP(0.5,Mx!A1:A6,Mx!B1:B6,,-1)', e('NA')],
  // the linear next value past an error and a blank (1, #N/A, blank, 3), a text key over numbers
  ['XLOOKUP', 'XLOOKUP(2,Er!A1:A4,Er!B1:B4,,1)', s('e4')],
  ['XLOOKUP', 'XLOOKUP(2,Er!A1:A4,Er!B1:B4,,-1)', s('e1')],
  ['XLOOKUP', 'XLOOKUP("x",Er!A1:A4,Er!B1:B4,,-1)', s('e4')],
  // a whole column searched from the end (9, 8, 7, 6, 5 and blanks)
  ['XLOOKUP', 'XLOOKUP(7.1,Dc!A:A,Dc!B:B,,-1,-2)', s('d3')],
  ['XLOOKUP', 'XLOOKUP(7.1,Dc!A:A,Dc!B:B,,1,-2)', s('d2')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): the binary search compares at the middle of
 * the positions it still has, both ends included, and an equal value there is found (5 among 1, 9, 5, 2, 3); a blank
 * counts as greater than the key and the positions run to the reference's end, so over 1, 2, blank, 4, 5, 6 the
 * search turns at the blank and misses 5, and a whole column searched from the end runs past its last value (the
 * next smaller #N/A, the next larger the last row). The set `xlookup-ironcalc-held-2` measures the rule they led to.
 */
export const MISSED: readonly string[] = ['XLOOKUP#16', 'XLOOKUP#19', 'XLOOKUP#33', 'XLOOKUP#34']
