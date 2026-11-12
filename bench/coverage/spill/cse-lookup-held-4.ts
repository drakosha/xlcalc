/**
 * The fourth held-out set of lookups in array formulas of the old kind (`run.ts spill cse-lookup-held-4`): the
 * dereferencing idioms of formulas written before dynamic arrays, `N(IF(1,…))` around INDEX's or XLOOKUP's selector and
 * `N(OFFSET(…))`, which 0.89.1 selects by position. Predicted before Excel computed it.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-lookup-held-4.predicted.js'

const held = (id: string, formula: string, area: string): SpillCase => {
  const at = area.split(':')[0]
  return one(id, `${formula} over ${area}`, { ...DATA, [at]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })
}

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

export function buildSpillCases(): SpillCase[] {
  return [
    held('index-nif-single', 'SUM(INDEX(B1:B3,N(IF(1,{2;3}))))', 'G1'),
    held('index-nif-two', 'SUM(INDEX(B1:B3,N(IF(1,{2;3}))))', 'G1:G2'),
    held('index-tif-single', 'COUNTA(INDEX(B1:B3,N(IF(1,{2;3}))))', 'G1'),
    held('index-if-single', 'SUM(INDEX(B1:B3,IF(1,{2;3})))', 'G1'),
    held('index-n-single', 'SUM(INDEX(B1:B3,N({2;3})))', 'G1'),
    held('index-nif-column-single', 'SUM(INDEX(B1:C3,N(IF(1,{2;3})),1))', 'G1'),
    held('offset-n-single', 'SUM(N(OFFSET(B1,{1;2},0)))', 'G1'),
    held('offset-n-two', 'SUM(N(OFFSET(B1,{1;2},0)))', 'G1:G2'),
    held('offset-nif-single', 'SUM(OFFSET(B1,N(IF(1,{1;2})),0))', 'G1'),
    held('xlookup-nif-single', 'SUM(XLOOKUP(N(IF(1,{1;2})),A1:A3,B1:B3))', 'G1'),
    held('index-nif-sequence-single', 'SUM(INDEX(B1:B3,N(IF(1,SEQUENCE(2)+1))))', 'G1'),
    held('vlookup-nif-single', 'SUM(VLOOKUP(N(IF(1,{1;2})),A1:B3,2,FALSE))', 'G1'),
  ]
}
