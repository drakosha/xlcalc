/**
 * Held out (`run.ts spill ref-arrays-held-2`): an array of references, OFFSET or INDIRECT over an array, given to the
 * functions the set `ref-arrays-held` left out. The arguments [MS-XLSX] 2.2.2 fills with a reference only
 * (`ref-argument-expression`, `xlsx/write/excel-loads.ts`): AGGREGATE's references, RANK's, AREAS, ROW, COLUMN, CELL's,
 * FORMULATEXT, ISFORMULA, a D function's database and criteria. And, for contrast, arguments that take an array as well
 * (AGGREGATE's array form, ROWS, INDEX, ISREF). Predicted before Excel computed it, by the rule the sets
 * `ref-arrays-held` and `indirect` measured: a function given an array of references where it takes a reference only
 * is called once for each reference and answers an array of the references' shape; where it takes an array, it takes
 * the array of the references' top-left values, as SUM does (`SUM(OFFSET(R!A1,{0;1},0,2))` is 30).
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './ref-arrays-held-2.predicted.js'

/**
 * Beside `DATA`: formulas for FORMULATEXT and ISFORMULA in E1:E3, a database J1:K4 (k: a, b, a; v: 1, 2, 4) and its
 * criteria L1:M2 (k = a, v > 1).
 */
const MORE: Record<string, CellSpec> = {
  E1: '=B1*2', E2: '=B2+1', E3: 5,
  J1: 'k', K1: 'v', J2: 'a', K2: 1, J3: 'b', K3: 2, J4: 'a', K4: 4,
  L1: 'k', M1: 'v', L2: 'a', M2: '>1',
}

/** A dynamic formula at G5 read over `read`. */
const dyn = (id: string, formula: string, read: string): SpillCase =>
  one(id, `${formula} at G5`, { ...DATA, ...MORE, G5: `=${formula}` }, [read], { predicted: PREDICTED[id] })

/** A legacy array formula over `area`. */
const cse = (id: string, formula: string, area: string): SpillCase =>
  one(id, `${formula} over ${area}`, { ...DATA, ...MORE, [area.split(':')[0]]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = []

export function buildSpillCases(): SpillCase[] {
  return [
    dyn('aggregate-9', 'AGGREGATE(9,0,OFFSET(B1,{0;1},0,2))', 'G5:G7'),
    dyn('aggregate-4-row', 'AGGREGATE(4,0,OFFSET(B1,0,{0,1},3))', 'G5:I5'),
    dyn('aggregate-3-indirect', 'AGGREGATE(3,5,INDIRECT({"B1:B2";"C1:C3"}))', 'G5:G7'),
    dyn('aggregate-two-arrays', 'AGGREGATE(9,0,OFFSET(B1,{0;1},0),OFFSET(C1,{0;1},0))', 'G5:G7'),
    dyn('aggregate-14-array-form', 'AGGREGATE(14,0,OFFSET(B1,{0;1},0,2),1)', 'G5:G7'),
    dyn('sum-aggregate', 'SUM(AGGREGATE(9,0,OFFSET(B1,{0;1;2},0)))', 'G5:G7'),
    dyn('rank', 'RANK(20,OFFSET(B1,{0;1},0,2))', 'G5:G7'),
    dyn('rank-avg-ascending', 'RANK.AVG(20,OFFSET(B1,0,0,{2;3}),1)', 'G5:G7'),
    dyn('areas', 'AREAS(OFFSET(B1,{0;1},0))', 'G5:G7'),
    dyn('row', 'ROW(OFFSET(B1,{0;1;2},0))', 'G5:G8'),
    dyn('column', 'COLUMN(OFFSET(B1,0,{0,1}))', 'G5:I5'),
    dyn('cell-address', 'CELL("address",OFFSET(B1,{0;1},0))', 'G5:G7'),
    dyn('formulatext', 'FORMULATEXT(OFFSET(E1,{0;1},0))', 'G5:G7'),
    dyn('isformula', 'ISFORMULA(OFFSET(E1,{0;1;2},0))', 'G5:G8'),
    dyn('dsum-database', 'DSUM(OFFSET(J1,0,0,{3;4},2),"v",L1:L2)', 'G5:G7'),
    dyn('dcount-criteria', 'DCOUNT(J1:K4,"v",OFFSET(L1,0,0,2,{1,2}))', 'G5:I5'),
    dyn('rows-values', 'ROWS(OFFSET(B1,{0;1},0,{1;2}))', 'G5:G7'),
    dyn('index-values', 'INDEX(OFFSET(B1,{0;1},0,2),2)', 'G5:G7'),
    dyn('isref', 'ISREF(OFFSET(B1,{0;1},0))', 'G5:G7'),
    cse('cse-aggregate', 'AGGREGATE(9,0,OFFSET(B1,{0;1},0,2))', 'G1:G2'),
    cse('cse-sum-rank', 'SUM(RANK(10,OFFSET(B1,0,0,{2;3})))', 'G1'),
  ]
}
