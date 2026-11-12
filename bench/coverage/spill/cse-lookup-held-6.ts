/**
 * The sixth held-out set of lookups in array formulas of the old kind (`run.ts spill cse-lookup-held-6`), a wider sweep
 * of idioms written before dynamic arrays: dereferencing variants, OFFSET and INDIRECT feeding functions that take a
 * reference, extraction with INDEX over SMALL(IF(…)), multi-criteria INDEX/MATCH, INDEX over arrays of columns, and ROW or
 * ROWS over the formula's own cell. Predicted before Excel computed it.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-lookup-held-6.predicted.js'

const held = (id: string, formula: string, area: string): SpillCase => {
  const at = area.split(':')[0]
  return one(id, `${formula} over ${area}`, { ...DATA, [at]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })
}

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

export function buildSpillCases(): SpillCase[] {
  return [
    held('nif-true-single', 'SUM(INDEX(B1:B3,N(IF(TRUE,{2;3}))))', 'G1'),
    held('nif-row-single', 'SUM(INDEX(B1:B3,N(IF(1,ROW(B2:B3)))))', 'G1'),
    held('max-nif-single', 'MAX(INDEX(B1:B3,N(IF(1,{1;3}))))', 'G1'),
    held('average-n-offset-single', 'AVERAGE(N(OFFSET(B1,{0;1;2},0)))', 'G1'),
    held('n-offset-two', 'SUM(N(OFFSET(B1,{0;1;2},0)))', 'G1:G2'),
    held('subtotal-offset-rows-two', 'SUM(SUBTOTAL(9,OFFSET(B1,ROW(B1:B3)-ROW(B1),0)))', 'G1:G2'),
    held('countif-offset-single', 'SUM(COUNTIF(OFFSET(A1,{0;1;2},0),">1"))', 'G1'),
    held('sumif-indirect-single', 'SUM(SUMIF(INDIRECT({"B1:B2";"B2:B3"}),">15"))', 'G1'),
    held('indirect-concat-single', 'SUM(INDIRECT("B"&{1;2}))', 'G1'),
    held('n-indirect-concat-single', 'SUM(N(INDIRECT("B"&{1;2})))', 'G1'),
    held('large-row-indirect-single', 'SUM(LARGE(B1:B3,ROW(INDIRECT("1:2"))))', 'G1'),
    held('small-if-constant-two', 'INDEX(B1:B3,SMALL(IF(A1:A3>1,ROW(A1:A3)),{1;2}))', 'G1:G2'),
    held('small-if-iferror-three', 'IFERROR(INDEX(B1:B3,SMALL(IF(A1:A3>1,ROW(A1:A3)),{1;2;3})),-1)', 'G1:G3'),
    held('small-if-rows-single', 'INDEX(B1:B3,SMALL(IF(A1:A3>1,ROW(A1:A3)),ROWS(E$1:E1)))', 'G1'),
    held('index-match-keys-two', 'INDEX(B1:B3,MATCH({1;2},A1:A3,0))', 'G1:G2'),
    held('index-match-criteria-single', 'INDEX(B1:B3,MATCH(1,(A1:A3=2)*(B1:B3>10),0))', 'G1'),
    held('sum-index-columns-single', 'SUM(INDEX(B1:C3,2,{1,2}))', 'G1'),
    held('sum-index-row-zero-single', 'SUM(INDEX(B1:C3,{1;2},0))', 'G1'),
    held('row-own-single', 'ROW(G1)', 'G1'),
    held('rows-own-single', 'ROWS(1:2)', 'G1'),
    held('row-own-two', 'ROW(1:2)', 'G1:G2'),
    held('sum-row-own-single', 'SUM(ROW(1:3))', 'G1'),
  ]
}
