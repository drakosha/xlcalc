/**
 * The fifth held-out set of lookups in array formulas of the old kind (`run.ts spill cse-lookup-held-5`): the
 * idioms of formulas written before dynamic arrays that take OFFSET's or INDIRECT's references whole: SUBTOTAL and SUMIF
 * over OFFSET, COUNTIF over INDIRECT, N over INDIRECT, and INDEX over SMALL(IF(…)) for extraction. Predicted before Excel
 * computed it.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-lookup-held-5.predicted.js'

const held = (id: string, formula: string, area: string): SpillCase => {
  const at = area.split(':')[0]
  return one(id, `${formula} over ${area}`, { ...DATA, [at]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })
}

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

export function buildSpillCases(): SpillCase[] {
  return [
    held('subtotal-offset-single', 'SUM(SUBTOTAL(9,OFFSET(B1,{0;1;2},0)))', 'G1'),
    held('subtotal-offset-two', 'SUM(SUBTOTAL(9,OFFSET(B1,{0;1;2},0)))', 'G1:G2'),
    held('subtotal-offset-rows-single', 'SUM(SUBTOTAL(3,OFFSET(B1,ROW(B1:B3)-ROW(B1),0)))', 'G1'),
    held('max-subtotal-offset-single', 'MAX(SUBTOTAL(9,OFFSET(B1,{0;1;2},0)))', 'G1'),
    held('sumif-offset-single', 'SUM(SUMIF(OFFSET(B1,{0;1},0,2),">15"))', 'G1'),
    held('countif-indirect-single', 'SUM(COUNTIF(INDIRECT({"B1:B2","B2:B3"}),">15"))', 'G1'),
    held('n-indirect-single', 'SUM(N(INDIRECT({"B2";"B3"})))', 'G1'),
    held('index-row-single', 'SUM(INDEX(B1:B3,ROW(B2:B3)))', 'G1'),
    held('index-small-if-two', 'INDEX(B1:B3,SMALL(IF(A1:A3>1,ROW(A1:A3)),ROW(1:2)))', 'G1:G2'),
  ]
}
