/**
 * Formulas inside the ranges the reference operators make, which a formula must read after they are computed: between
 * the ends of the range operator (static ends, ends from INDEX, a spill reference as an end, `S8:S9!A1` on its own
 * sheet), in an intersection, on the middle sheet of a reference across sheets, in the second area of a union that
 * SUBTOTAL, AGGREGATE and FREQUENCY take. Answers: `test/excel/cases/ref-deps.json`; `PREDICTED` was committed
 * before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** R: values and formulas between them (B1 50, D1 8, C3 5, B5 spills 1, 2, 3); S9: a formula over `S8:S9!A1`, which is A1:S8 there; Uno, Due, Tre: a formula on the middle one. */
export const SHEETS: SheetCells = {
  R: { A1: 1, B1: { formula: 'E1*10' }, C1: 1, D1: { formula: 'E9*2' }, E1: 5, G1: 3, C3: { formula: 'E9+1' }, B5: { formula: '_xlfn.SEQUENCE(1,3)', dynamic: true }, E9: 4 },
  S9: { A1: 7, B1: { formula: 'E9*10' }, E9: 4, C9: { formula: 'SUM(S8:S9!A1)' } },
  Uno: { A1: 1 },
  Due: { A1: { formula: 'R!E9*3' } },
  Tre: { A1: 3 },
}

const n = (v: number): Val => ({ t: 'n', v })
const s = (v: string): Val => ({ t: 's', v })

/** Group, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['range', 'SUM(R!A1:R!A1:R!C1)', n(52)],
  ['range', 'SUM(INDEX(R!A1:C1,1,1):INDEX(R!E1:G1,1,3))', n(68)],
  ['range', 'SUM(R!A1:R!B5#)', n(71)],
  ['range', 'ROWS(R!A1:R!B5#)', n(5)],
  ['range', "'S9'!C9", n(47)],
  ['range', 'SUM(R!A1:INDEX(R!A1:D3,3,4))', n(65)],
  ['space', 'SUM(R!A1:D1 R!B1:B2)', n(50)],
  ['cube', 'SUM(Uno:Tre!A1)', n(16)],
  ['union', 'SUBTOTAL(9,(R!E1,R!B1:D1))', n(64)],
  ['union', 'AGGREGATE(9,4,(R!E1,R!B1:D1))', n(64)],
  ['union', 'ARRAYTOTEXT(FREQUENCY((R!E1,R!B1:D1),{10;100}),1)', s('{3;1;0}')],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, f] of CASES) add(group, group, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][2]]))
