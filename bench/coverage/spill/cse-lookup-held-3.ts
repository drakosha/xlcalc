/**
 * The third held-out set of lookups in array formulas of the old kind (`run.ts spill cse-lookup-held-3`): the rules
 * `cse-lookup-held-2` led to beyond what it measured. Paired arrays other than SUMX2PY2 (CORREL, SLOPE, FORECAST),
 * MODE.SNGL, SUMPRODUCT over INDEX, AVERAGE over OFFSET, HLOOKUP and CHOOSEROWS. Predicted before Excel computed it.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-lookup-held-3.predicted.js'

const held = (id: string, formula: string, area: string): SpillCase => {
  const at = area.split(':')[0]
  return one(id, `${formula} over ${area}`, { ...DATA, [at]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })
}

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

const THREE = 'XLOOKUP({1;2;3},A1:A3,B1:B3)'

export function buildSpillCases(): SpillCase[] {
  return [
    held('correl-two', `CORREL(${THREE},{1;2;3})`, 'G1:G2'),
    held('slope-two', `SLOPE(${THREE},{1;2;3})`, 'G1:G2'),
    held('forecast-two', `FORECAST(4,${THREE},{1;2;3})`, 'G1:G2'),
    held('mode-sngl-two', 'MODE.SNGL(XLOOKUP({1;2},A1:A3,B1:B3),10)', 'G1:G2'),
    held('sumproduct-index-two', 'SUMPRODUCT(INDEX(B1:B3,{2;3}))', 'G1:G2'),
    held('sumproduct-index-single', 'SUMPRODUCT(INDEX(B1:B3,{2;3}))', 'G1'),
    held('average-offset-two', 'AVERAGE(OFFSET(B1,{1;2},0))', 'G1:G2'),
    held('hlookup-two', 'SUM(HLOOKUP({30,300},A1:C2,2,FALSE))', 'G1:H1'),
    held('chooserows-two', 'SUM(CHOOSEROWS(B1:B3,{2;3}))', 'G1:G2'),
  ]
}
