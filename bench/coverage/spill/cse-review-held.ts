/**
 * Held out (`run.ts spill cse-review-held`), after a review found two faults: arithmetic around an aggregation in a
 * legacy array formula over several cells (`SUM(XLOOKUP({1;2},…))+1`), which the engine let take the aggregation's
 * inside whole, and a function lifted over an array of values and an array of references at once
 * (`COUNTIF(INDIRECT({"B1";"B2:B3"}),{">0";">0"})`), which kept the first answer of each reference. Predicted before
 * Excel computed it: the argument of a function that answers one value selects by position whatever is around the call;
 * arrays of values and of references lift a call together, as arrays do.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-review-held.predicted.js'

const LOOK = 'XLOOKUP({1;2},A1:A3,B1:B3)'

const dyn = (id: string, formula: string, at: string, read: string): SpillCase =>
  one(id, `${formula} at ${at}`, { ...DATA, [at]: `=${formula}` } as Record<string, CellSpec>, [read], { predicted: PREDICTED[id] })

const cse = (id: string, formula: string, area: string): SpillCase =>
  one(id, `${formula} over ${area}`, { ...DATA, [area.split(':')[0]]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })

/**
 * The cases whose prediction Excel answered otherwise (Excel for Mac and Windows alike): TRANSPOSE evaluates its argument
 * as an array in a legacy array formula over several cells, SUM or not around it, 31; 31.
 */
export const MISSED: readonly string[] = ['sum-transpose-plus']

export function buildSpillCases(): SpillCase[] {
  return [
    cse('sum-plus-one', `SUM(${LOOK})+1`, 'G1:G2'),
    cse('abs-sum', `ABS(SUM(${LOOK}))`, 'G1:G2'),
    cse('sum-plus-sum', `SUM(${LOOK})+SUM(${LOOK})`, 'G1:G2'),
    cse('max-index-times', 'MAX(INDEX(B1:B3,{2;3}))*2', 'G1:G2'),
    cse('sum-offset-minus', 'SUM(OFFSET(B1,{1;2},0))-1', 'G1:G2'),
    cse('average-plus-sum', `AVERAGE(${LOOK})+SUM(B1:B3)`, 'G1:G2'),
    cse('sum-transpose-plus', `SUM(TRANSPOSE(${LOOK}))+1`, 'G1:G2'),
    cse('if-sum-plus', `IF(TRUE,SUM(${LOOK}))+0`, 'G1:G2'),
    cse('sum-plus-one-single', `SUM(${LOOK})+1`, 'G1'),
    cse('sum-inner-plus', `SUM(${LOOK}+1)`, 'G1:G2'),
    dyn('countif-two-arrays', 'COUNTIF(INDIRECT({"B1";"B2:B3"}),{">0";">0"})', 'G5', 'G5:G6'),
    dyn('subtotal-function-array', 'SUBTOTAL({9;1},OFFSET(B1,{0;1},0,2))', 'G5', 'G5:G6'),
    dyn('countif-criteria-row', 'COUNTIF(INDIRECT({"B1:B2";"B2:B3"}),{">15",">5"})', 'G5', 'G5:H6'),
    dyn('sumif-criteria-array', 'SUMIF(OFFSET(B1,{0;1},0,2),{">15";">5"})', 'G5', 'G5:G6'),
    dyn('subtotal-function-longer', 'SUBTOTAL({9;1;2},OFFSET(B1,{0;1},0,2))', 'G5', 'G5:G7'),
    cse('cse-countif-two-arrays', 'COUNTIF(INDIRECT({"B1";"B2:B3"}),{">0";">0"})', 'G1:G2'),
  ]
}
