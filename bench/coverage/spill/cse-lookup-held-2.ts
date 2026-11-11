/**
 * The second held-out set of lookups in array formulas of the old kind (`run.ts spill cse-lookup-held-2`), after
 * `cse-lookup-held` showed SUMPRODUCT evaluating its argument whole over two cells and OFFSET selecting by position:
 * other arguments of Excel's array class, and the other reference INDIRECT finds. Predicted before Excel computed it.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-lookup-held-2.predicted.js'

const held = (id: string, formula: string, area: string): SpillCase => {
  const at = area.split(':')[0]
  return one(id, `${formula} over ${area}`, { ...DATA, [at]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })
}

const KEYS = 'XLOOKUP({1;2},A1:A3,B1:B3)'

export function buildSpillCases(): SpillCase[] {
  return [
    held('rows-two', `ROWS(${KEYS})`, 'G1:G2'),
    held('mode-two', `MODE(${KEYS},10)`, 'G1:G2'),
    held('sumx2py2-two', `SUMX2PY2(${KEYS},{0;0})`, 'G1:G2'),
    held('indirect-two', 'SUM(INDIRECT({"B2";"B3"}))', 'G1:G2'),
    held('indirect-single', 'SUM(INDIRECT({"B2";"B3"}))', 'G1'),
    held('offset-height-two', 'SUM(OFFSET(B1,0,0,{1;2}))', 'G1:G2'),
  ]
}
