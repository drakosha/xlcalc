/**
 * Held out: low-severity edges from the review of 21.10. SUMIF and AVERAGEIF whose sum range, resized to the criteria
 * range, runs past the sheet's last row or column; the percent operator on a value whose hundredth is subnormal (as `/`
 * gives 0 there); TEXT with an elapsed code ([h], [mm], [ss]) whose total passes 1E+21; and UNIQUE over text that
 * compares equal by `=` but differs in its code points (é in NFC and NFD, the ligature ﬁ, a soft hyphen). The values
 * of sheet V: A1:A3 = 1, 2, 3, B1048575:B1048576 = 10, 20, H1:J1 = 1, 2, 3, XFC1:XFD1 = 1000, 2000, K1:K3 = 1E-307,
 * 2E-306, 3E-306, M1:M4 = é, e + U+0301, a, a + U+00AD. Answers: `test/excel/cases/low-values-held.json`; `PREDICTED`,
 * the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldGroups } from './arrays/common.js'
import { b, e, n, s } from '../vals.js'

export const SHEETS: SheetCells = {
  V: {
    A1: 1, A2: 2, A3: 3, B1048575: 10, B1048576: 20, H1: 1, I1: 2, J1: 3, XFC1: 1000, XFD1: 2000,
    K1: 1e-307, K2: 2e-306, K3: 3e-306,
    M1: 'é', M2: 'é', M3: 'a', M4: 'a­',
  },
}

/** Group, function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, string, Val][] = [
  // the sum range past the sheet's edge: the cells that exist, positions past it blank
  ['sumif-edge', 'SUMIF', 'SUMIF(V!A1:A3,">0",V!B1048575)', n(30)],
  ['sumif-edge', 'SUMIF', 'SUMIF(V!A1:A3,">2",V!B1048575)', n(0)],
  ['sumif-edge', 'SUMIF', 'SUMIF(V!A1:A3,">1",V!B1048575)', n(20)],
  ['sumif-edge', 'SUMIF', 'SUMIF(V!A1:A3,">0",V!B1048576)', n(20)],
  ['sumif-edge', 'SUMIF', 'SUMIF(V!H1:J1,">0",V!XFC1)', n(3000)],
  ['sumif-edge', 'SUMIF', 'SUMIF(V!A1:A3,">0",V!B1048574)', n(30)],
  ['sumif-edge', 'AVERAGEIF', 'AVERAGEIF(V!A1:A3,">0",V!B1048575)', n(15)],
  ['sumif-edge', 'AVERAGEIF', 'AVERAGEIF(V!A1:A3,">2",V!B1048575)', e('DIV0')],
  // a hundredth below the smallest normal number is 0, as `/` gives it (probes text#545, text#546); 3E-308 is normal
  ['percent', 'operators', '"a"&(V!K1%)', s('a0')],
  ['percent', 'operators', 'V!K1%=0', b(true)],
  ['percent', 'operators', 'V!K2%*1E300', n(0)],
  ['percent', 'operators', '-V!K1%=0', b(true)],
  ['percent', 'operators', '"a"&(V!K3%)', s('a3E-308')],
  ['percent', 'operators', '"a"&(V!K1/100)', s('a0')],
  // an elapsed total of 1E+21 and more: digits, 15 significant and zeros
  ['elapsed', 'TEXT', 'TEXT(1E+20,"[h]")', s('2400000000000000000000')],
  ['elapsed', 'TEXT', 'TEXT(1E+17,"[ss]")', s('8640000000000000000000')],
  ['elapsed', 'TEXT', 'TEXT(-1E+20,"[mm]")', s('-144000000000000000000000')],
  ['elapsed', 'TEXT', 'TEXT(4.2E+19,"[h]")', s('1008000000000000000000')],
  ['elapsed', 'TEXT', 'TEXT(1E+16,"[h]")', s('240000000000000000')],
  ['elapsed', 'TEXT', 'LEN(TEXT(1E+250,"[h]"))', n(252)],
  ['elapsed', 'TEXT', 'TEXT(1E+300,"[h]")', e('VALUE')],
  // text equal by `=` and MATCH, distinct by code points: what UNIQUE takes as one
  ['unique-text', 'UNIQUE', 'ROWS(UNIQUE(V!M1:M2))', n(2)],
  ['unique-text', 'UNIQUE', 'ROWS(UNIQUE(V!M1:M2,FALSE,TRUE))', n(2)],
  ['unique-text', 'UNIQUE', 'ROWS(UNIQUE({"ﬁ";"fi"}))', n(2)],
  ['unique-text', 'UNIQUE', 'ROWS(UNIQUE(V!M3:M4))', n(2)],
  ['unique-text', 'UNIQUE', 'ROWS(UNIQUE({"a";"A"}))', n(1)],
  ['unique-text', '=', 'V!M1=V!M2', b(true)],
  ['unique-text', '=', '"ﬁ"="fi"', b(true)],
  ['unique-text', '=', 'V!M3=V!M4', b(true)],
  ['unique-text', 'MATCH', 'MATCH(V!M2,V!M1:M2,0)', n(1)],
  ['unique-text', 'COUNTIF', 'COUNTIF(V!M1:M2,V!M1)', n(1)],
  ['unique-text', 'EXACT', 'EXACT(V!M1,V!M2)', b(false)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldGroups(CASES)

/**
 * The predictions Excel did not confirm, the same in Excel for Mac and for Windows: UNIQUE takes é in NFC and NFD, and ﬁ
 * and fi, as one (the engine then followed); `=` tells a soft hyphen apart, MATCH NFC from NFD, and COUNTIF counts them
 * as one (`textEquivalence` in KNOWN).
 */
export const MISSED = ['unique-text#1', 'unique-text#2', 'unique-text#3', 'unique-text#8', 'unique-text#9', 'unique-text#10']
