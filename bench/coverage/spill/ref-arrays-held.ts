/**
 * Held out (`run.ts spill ref-arrays-held`): an array of references, OFFSET or INDIRECT over an array, given to a
 * function that takes a reference, in dynamic formulas and in legacy array formulas: AVERAGEIF with and without its
 * average range, SUMIFS, COUNTIFS and MAXIFS over two such arrays at once, COUNTBLANK, SUBTOTAL 1, 2, 4 and 9 (two
 * arrays of references in one call, a row of them, a two-dimensional one, one with an error), INDIRECT with the sheet in
 * its text. Predicted before Excel computed it, by the rule the sets `cse-lookup-held-5` and `own-place-held` measured:
 * the function is called once for each reference and answers an array of the references' shape.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './ref-arrays-held.predicted.js'

/** A dynamic formula at `at` read over `read`. */
const dyn = (id: string, formula: string, at: string, read: string): SpillCase =>
  one(id, `${formula} at ${at}`, { ...DATA, [at]: `=${formula}` } as Record<string, CellSpec>, [read], { predicted: PREDICTED[id] })

/** A legacy array formula over `area`. */
const cse = (id: string, formula: string, area: string): SpillCase =>
  one(id, `${formula} over ${area}`, { ...DATA, [area.split(':')[0]]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

export function buildSpillCases(): SpillCase[] {
  return [
    dyn('averageif-offset', 'AVERAGEIF(OFFSET(B1,{0;1},0,2),">15")', 'G5', 'G5:G6'),
    dyn('averageif-offset-both', 'AVERAGEIF(OFFSET(B1,{0;1},0,2),">15",OFFSET(C1,{0;1},0,2))', 'G5', 'G5:G6'),
    dyn('sumifs-two', 'SUMIFS(OFFSET(C1,{0;1},0,2),OFFSET(B1,{0;1},0,2),">15")', 'G5', 'G5:G6'),
    dyn('sumifs-fixed-sum', 'SUMIFS(C1:C2,OFFSET(B1,{0;1},0,2),">15")', 'G5', 'G5:G6'),
    dyn('countifs-two', 'COUNTIFS(OFFSET(B1,{0;1},0,2),">15",OFFSET(C1,{0;1},0,2),">150")', 'G5', 'G5:G6'),
    dyn('maxifs-two', 'MAXIFS(OFFSET(C1,{0;1},0,2),OFFSET(B1,{0;1},0,2),">5")', 'G5', 'G5:G6'),
    dyn('countblank-offset', 'COUNTBLANK(OFFSET(B3,{0;1},0,2))', 'G5', 'G5:G6'),
    dyn('subtotal-1', 'SUBTOTAL(1,OFFSET(B1,{0;1},0,2))', 'G5', 'G5:G6'),
    dyn('subtotal-4-row', 'SUBTOTAL(4,OFFSET(B1,0,{0,1},3))', 'G5', 'G5:H5'),
    dyn('subtotal-2-indirect', 'SUBTOTAL(2,INDIRECT({"B1:B2";"C1:C3"}))', 'G5', 'G5:G6'),
    dyn('subtotal-indirect-sheet', 'SUBTOTAL(9,INDIRECT({"S!B1";"S!C1:C2"}))', 'G5', 'G5:G6'),
    dyn('subtotal-2d', 'SUBTOTAL(9,OFFSET(B1,{0,1},{0;1}))', 'G5', 'G5:H6'),
    dyn('subtotal-ref-error', 'SUBTOTAL(9,OFFSET(B1,{0;-5},0))', 'G5', 'G5:G6'),
    dyn('subtotal-two-arrays', 'SUBTOTAL(9,OFFSET(B1,{0;1},0),OFFSET(C1,{0;1},0))', 'G5', 'G5:G6'),
    dyn('sum-countif-offset', 'SUM(COUNTIF(OFFSET(B1,{0;1;2},0),">15"))', 'G5', 'G5'),
    dyn('sumproduct-subtotal', 'SUMPRODUCT(SUBTOTAL(3,OFFSET(B1,{0;1;2},0)))', 'G5', 'G5'),
    dyn('n-offset', 'N(OFFSET(B1,{0;1},0))', 'G5', 'G5:G6'),
    cse('cse-subtotal-three', 'SUBTOTAL(9,OFFSET(B1,{0;1;2},0))', 'G1:G3'),
    cse('cse-sumif-two', 'SUMIF(OFFSET(B1,{0;1},0,2),">15")', 'G1:G2'),
    cse('cse-sum-countif-two', 'SUM(COUNTIF(INDIRECT({"B1:B2";"B2:B3"}),">15"))', 'G1:G2'),
    cse('cse-averageif-single', 'SUM(AVERAGEIF(OFFSET(B1,{0;1},0,2),">15"))', 'G1'),
  ]
}
