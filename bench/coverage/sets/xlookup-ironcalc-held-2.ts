/**
 * XLOOKUP's and XMATCH's binary search as `xlookup-ironcalc-held`'s misses showed it, on new data: unsorted values
 * where the middle finds or passes the key, a blank or an error at the middle, runs of equal values ascending and
 * descending, a reference longer than its values, whole columns ascending and descending past their last value.
 * Answers: `test/excel/cases/xlookup-ironcalc-held-2.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

/** Column A of a sheet holds `values` (null a blank), column B the text `<sheet><row>` beside each. */
const data = (name: string, values: (number | null | { formula: string })[]) =>
  Object.fromEntries(values.flatMap((v, i) => [...(v === null ? [] : [[`A${i + 1}`, v]]), [`B${i + 1}`, `${name.toLowerCase()}${i + 1}`]]))

export const SHEETS: SheetCells = {
  Us: data('Us', [3, 8, 1, 9, 5, 2, 7]),
  Bm: data('Bm', [1, 2, 3, null, 5, 6, 7]),
  Em: data('Em', [1, 2, 3, { formula: 'NA()' }, 5, 6, 7]),
  Dd: data('Dd', [9, 8, null, 6, 5]),
  Ru: data('Ru', [1, 3, 3, 3, 3, 5, 7]),
  Rd: data('Rd', [7, 5, 3, 3, 3, 3, 1]),
  Ab: data('Ab', [1, 2, 3, 4, 5, null, null, null, null, null]),
  Aw: { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: 'aw1', B2: 'aw2', B3: 'aw3', B4: 'aw4', B5: 'aw5' },
  Dw: { A1: 9, A2: 8, A3: 7, A4: 6, A5: 5, B1: 'dw1', B2: 'dw2', B3: 'dw3', B4: 'dw4', B5: 'dw5' },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // unsorted 3, 8, 1, 9, 5, 2, 7: the middle passes 5 and 1, and finds 8 on the way
  ['XLOOKUP', 'XLOOKUP(5,Us!A1:A7,Us!B1:B7,,0,2)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(8,Us!A1:A7,Us!B1:B7,,0,2)', s('us2')],
  ['XLOOKUP', 'XLOOKUP(1,Us!A1:A7,Us!B1:B7,,0,2)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(5,Us!A1:A7,Us!B1:B7,,-1,2)', s('us1')],
  ['XMATCH', 'XMATCH(5,Us!A1:A7,1,2)', n(2)],
  // a blank at the middle (1, 2, 3, blank, 5, 6, 7) and an error there
  ['XLOOKUP', 'XLOOKUP(6,Bm!A1:A7,Bm!B1:B7,,0,2)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(2,Bm!A1:A7,Bm!B1:B7,,0,2)', s('bm2')],
  ['XLOOKUP', 'XLOOKUP(6,Bm!A1:A7,Bm!B1:B7,,-1,2)', s('bm3')],
  ['XLOOKUP', 'XLOOKUP(6,Bm!A1:A7,Bm!B1:B7,,1,2)', s('bm4')],
  ['XLOOKUP', 'XLOOKUP(6,Em!A1:A7,Em!B1:B7,,0,2)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(2.5,Em!A1:A7,Em!B1:B7,,1,2)', s('em3')],
  // descending 9, 8, blank, 6, 5: the blank at the middle sends the search on to the end
  ['XLOOKUP', 'XLOOKUP(6,Dd!A1:A5,Dd!B1:B5,,0,-2)', s('dd4')],
  ['XLOOKUP', 'XLOOKUP(8,Dd!A1:A5,Dd!B1:B5,,0,-2)', e('NA')],
  // runs of equal values: 1, 3, 3, 3, 3, 5, 7 and 7, 5, 3, 3, 3, 3, 1
  ['XLOOKUP', 'XLOOKUP(3,Ru!A1:A7,Ru!B1:B7,,0,2)', s('ru2')],
  ['XLOOKUP', 'XLOOKUP(3,Ru!A1:A7,Ru!B1:B7,,1,2)', s('ru2')],
  ['XLOOKUP', 'XLOOKUP(4,Ru!A1:A7,Ru!B1:B7,,-1,2)', s('ru5')],
  ['XLOOKUP', 'XLOOKUP(3,Rd!A1:A7,Rd!B1:B7,,0,-2)', s('rd6')],
  ['XLOOKUP', 'XLOOKUP(4,Rd!A1:A7,Rd!B1:B7,,-1,-2)', s('rd3')],
  ['XLOOKUP', 'XLOOKUP(4,Rd!A1:A7,Rd!B1:B7,,1,-2)', s('rd2')],
  // a reference longer than its values (1 to 5, then five blanks)
  ['XLOOKUP', 'XLOOKUP(3,Ab!A1:A10,Ab!B1:B10,,0,2)', s('ab3')],
  ['XLOOKUP', 'XLOOKUP(10,Ab!A1:A10,Ab!B1:B10,,1,2)', s('ab6')],
  ['XLOOKUP', 'XLOOKUP(10,Ab!A1:A10,Ab!B1:B10,,-1,2)', s('ab5')],
  // whole columns: ascending 1 to 5 and descending 9 to 5, then blanks
  ['XLOOKUP', 'XLOOKUP(3,Aw!A:A,Aw!B:B,,0,2)', s('aw3')],
  ['XLOOKUP', 'XLOOKUP(10,Aw!A:A,Aw!B:B,,1,2)', n(0)],
  ['XMATCH', 'XMATCH(10,Aw!A:A,1,2)', n(6)],
  ['XLOOKUP', 'XLOOKUP(7,Dw!A:A,Dw!B:B,,0,-2)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(7,Dw!A1:A5,Dw!B1:B5,,0,-2)', s('dw3')],
  ['XMATCH', 'XMATCH(7.1,Dw!A:A,-1,-2)', e('NA')],
  ['XMATCH', 'XMATCH(7.1,Dw!A:A,1,-2)', n(1048576)],
  // wildcards with binary search in XMATCH
  ['XMATCH', 'XMATCH("a*",Ab!B1:B10,2,-2)', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
