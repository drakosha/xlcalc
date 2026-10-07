/**
 * Held-out formulas of the old kind (`run.ts spill legacy-held`): the engine computes each converted by `fromLegacy`,
 * and the answers predicted from the rules of the `legacy` sets were committed before Excel computed them
 * (`PREDICTED`).
 */
import { sheet, type CellSpec, type SpillCase } from './cases.js'
import { PREDICTED } from './legacy-held.predicted.js'

const DATA: Record<string, CellSpec> = {
  ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`A${i + 1}`, i + 1])),
  ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`B${i + 1}`, (i + 1) * 10])),
}

const FORMULAS: Record<string, string> = {
  'T!B4': '=S!A1:A10*S!B1:B10',
  'S!H6': '=IF(A1:A10>5,"big","small")',
  'S!H2': '=SUMPRODUCT((A1:A10>5)*B1:B10)',
  'S!H3': '=INDEX(A1:B10,0,1)*10',
  'S!H4': '=nm*2',
  'S!H5': '=ROW(A1:A10)*2',
  'S!H7': '={5,6,7}+A1:A10',
  'S!H8': '=SUM(A1:A10*2)',
  'S!H9': '=VLOOKUP(A:A,A:B,2,FALSE)',
  'S!H11': '=A1:A10',
  'T!E2': '=SUM(S!A:A)',
  'S!H12': '=TRANSPOSE(A1:A3)',
  'S!H10': '=LET(x,A1:A10,x*3)',
}

export function buildSpillCases(): SpillCase[] {
  const cells: Record<'S' | 'T', Record<string, CellSpec>> = { S: { ...DATA }, T: {} }
  for (const [at, f] of Object.entries(FORMULAS)) cells[at.slice(0, 1) as 'S' | 'T'][at.slice(2)] = f
  return [
    {
      id: 'positions',
      about: 'formulas of the old kind over both sheets, names, arrays and LET',
      book: { sheets: [sheet('S', cells.S), sheet('T', cells.T)], names: [{ name: 'nm', ref: "'S'!$B$1:$B$10" }] },
      read: Object.keys(FORMULAS).map((k) => ({ sheet: k.slice(0, 1), range: k.slice(2) })),
      legacy: Object.keys(FORMULAS),
      predicted: PREDICTED.positions,
    },
  ]
}
