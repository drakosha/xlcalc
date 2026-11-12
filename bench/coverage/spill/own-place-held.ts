/**
 * Held out (`run.ts spill own-place-held`): dynamic array formulas whose ROW, COLUMN or ROWS reads the place of the cells
 * they spill over (`=ROW(1:2)` in G1), which the engine took for a circular reference, and SUBTOTAL, SUMIF and COUNTIF
 * over an array of references from OFFSET or INDIRECT. Predicted before Excel computed it.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { DATA } from './cse-lookup.js'
import { PREDICTED } from './own-place-held.predicted.js'

const held = (id: string, formula: string, at: string, read: string): SpillCase =>
  one(id, `${formula} at ${at}`, { ...DATA, [at]: `=${formula}` } as Record<string, CellSpec>, [read], { predicted: PREDICTED[id] })

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

export function buildSpillCases(): SpillCase[] {
  return [
    held('row-rows', 'ROW(1:2)', 'G1', 'G1:G2'),
    held('row-range', 'ROW(G1:G3)', 'G1', 'G1:G3'),
    held('column-range', 'COLUMN(G1:I1)', 'G1', 'G1:I1'),
    held('rows-own', 'ROWS(G1:G5)+SEQUENCE(2)', 'G1', 'G1:G2'),
    held('subtotal-offset', 'SUBTOTAL(9,OFFSET(B1,{0;1;2},0))', 'G5', 'G5:G7'),
    held('sum-subtotal-offset', 'SUM(SUBTOTAL(9,OFFSET(B1,{0;1;2},0)))', 'G5', 'G5'),
    held('countif-indirect', 'COUNTIF(INDIRECT({"B1:B2";"B2:B3"}),">15")', 'G5', 'G5:G6'),
    held('sumif-offset', 'SUMIF(OFFSET(B1,{0;1},0,2),">15")', 'G5', 'G5:G6'),
  ]
}
