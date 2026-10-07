/**
 * Formulas of the old kind (`run.ts spill legacy`): Excel 365 computes them with implicit intersection, a range
 * where one value goes taking the cell of the formula's row or column, on the formula's sheet or another, and shows
 * them with `@` (`formula2`). Each formula's cell is listed in `legacy`; the engine computes the book with them
 * converted by `fromLegacy`.
 */
import { sheet, type CellSpec, type SpillCase } from './cases.js'
import type { NameInput } from '../../../src/contract.js'

/** Data on S: A1:A10 1…10, B1:B10 10…100, C1:E1 a row, a 2×2 block G1:H2. */
const DATA: Record<string, CellSpec> = {
  ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`A${i + 1}`, i + 1])),
  ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`B${i + 1}`, (i + 1) * 10])),
  C1: 7, D1: 8, E1: 9, G1: 1, H1: 2, G2: 3, H2: 4,
}

/** A case whose formulas (on S, or on T as `T!B7`) are all of the old kind, each cell read; `names` of the book. */
function legacy(id: string, about: string, formulas: Record<string, string>, names?: NameInput[]): SpillCase {
  const cells: Record<'S' | 'T', Record<string, CellSpec>> = { S: { ...DATA }, T: {} }
  const keys = Object.entries(formulas).map(([at, f]) => {
    const [s, a] = at.startsWith('T!') ? ['T' as const, at.slice(2)] : ['S' as const, at]
    cells[s][a] = f
    return { sheet: s, range: a }
  })
  return { id, about, book: { sheets: [sheet('S', cells.S), sheet('T', cells.T)], ...(names ? { names } : {}) }, read: keys, legacy: keys.map((k) => `${k.sheet}!${k.range}`) }
}

export function buildSpillCases(): SpillCase[] {
  return [
    legacy('column', "a column takes the formula's row, else #VALUE!", { J3: '=A1:A10', J12: '=A1:A10', J4: '=A1:A10*2', J5: '=A:A', J6: '=$A$1:$A$10+B1:B10', 'T!B7': '=S!A1:A10', 'T!B11': '=S!A1:A10' }),
    legacy('row', "a row takes the formula's column", { D12: '=C1:E1', I12: '=C1:E1', C13: '=1:1', 'T!D5': '=S!C1:E1' }),
    legacy('block', "a block takes the cell in the formula's row and column", { 'T!G2': '=S!G1:H2', 'T!H1': '=S!G1:H2', 'T!J2': '=S!G1:H2', 'T!G5': '=S!G1:H2', J1: '=G1:H2' }),
    legacy('calls', 'ranges in calls: one value taken in a value argument, the whole range in a reference or array argument', {
      K1: '=SUM(A1:A10)',
      K2: '=SUM(A1:A10*2)',
      K3: '=SUMPRODUCT(A1:A10*2)',
      K4: '=ABS(A1:A10)',
      K5: '=IF(A1:A10>4,"y","n")',
      K6: '=VLOOKUP(A1:A10,A1:B10,2,FALSE)',
      K7: '=MATCH(A1:A10,A1:A10,0)',
      K8: '=COUNTIF(A1:A10,A1:A10)',
      K9: '=SUMIF(A1:A10,">3",B1:B10)',
      K10: '=LEN(B1:B10)',
      K11: '=ROW(A1:A10)',
      K12: '=INDEX(A1:B10,0,2)',
      K13: '=INDEX(A1:B10,3,2)',
      K14: '=SUM(IF(A1:A10>4,1,0))',
      K15: '=SUMPRODUCT((A1:A10>4)*1)',
      K16: '=COLUMN(C1:E1)',
    }),
    legacy('arrays', 'arrays: a constant, a function with an array answer, a lookup over a whole column', {
      L1: '={1,2,3}',
      L2: '={1,2,3}*2',
      L3: '=SEQUENCE(3)',
      L4: '=TRANSPOSE(A1:A3)',
      L5: '=MMULT(G1:H2,G1:H2)',
      L6: '=VLOOKUP(A:A,A:B,2,FALSE)',
      L7: '=SUM(COUNTIF(A1:A10,{3,4}))',
      L8: '=INDEX(A1:A10*2,7)',
      L9: '=IF(TRUE,A1:A10)',
      L10: '=CHOOSE(1,A1:A10,B1:B10)',
    }),
    legacy(
      'names',
      'names over a range, a constant, an array, a formula, a union',
      { M2: '=nr', M3: '=nr*2', M4: '=SUM(nr)', M5: '=nc', M6: '=na', M7: '=nf', M8: '=nu', M9: '=INDEX(nr,2)', M10: '=SUM(nf)' },
      [
        { name: 'nr', ref: "'S'!$A$1:$A$10" },
        { name: 'nc', ref: '0.2' },
        { name: 'na', ref: '{4,5,6}' },
        { name: 'nf', ref: "'S'!$A$1:$A$10*3" },
        { name: 'nu', ref: "'S'!$A$1:$A$3,'S'!$B$1:$B$3" },
      ],
    ),
  ]
}
