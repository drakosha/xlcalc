/**
 * Held out after the set `indirect` was fitted: the rules it led to on other cases (OFFSET truncating its numbers and
 * taking a size left out as the reference's, CELL, SHEET, FORMULATEXT and ISFORMULA over a union's first area, the
 * last two cell by cell over a range, SHEETS and SHEET of values, TRIMRANGE's whole codes, the trim references), and,
 * in a book of their own, the cases Excel for Mac answered `#NUM!` in that set's run. Answers:
 * `test/excel/cases/indirect-held.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'
import { SHEETS } from './indirect.js'

export { SHEETS }

const CASES: Record<string, [fn: string, formulas: string[]]> = {
  again: ['INDIRECT', ['INDIRECT(D!A7)', 'SUM(INDIRECT({"R!A1","R!A2"}))', 'SUMPRODUCT(ROW(INDIRECT("1:4")))', 'OFFSET(INDIRECT("R!A1"),2,0)', 'CELL("row",INDIRECT("R!A5"))', 'FORMULATEXT(R!A1)', 'SUM(INDIRECT("R!E1#"))', 'ROWS(INDIRECT("R!E1#"))', 'INDEX(INDIRECT("R!E1#"),2)', 'SUM(INDIRECT(R!B5))']],
  offset: ['OFFSET', ['OFFSET(R!A1,0.999,0)', 'OFFSET(R!A2,-0.5,0)', 'ROWS(OFFSET(R!A1,0,0,1.9))', 'SUM(OFFSET(R!A2,0,0,-1.5))', 'SUM(OFFSET(R!A1,0,0,,2))', 'COLUMNS(OFFSET(R!A1:C1,1,0,,))', 'OFFSET(R!A1,"0.5",0)', 'SUM(OFFSET(R!A1:A2,1,0))']],
  cell: ['CELL', ['CELL("row",(R!B2,R!A1))', 'MID(CELL("address",R:C!B2),FIND("]",CELL("address",R:C!B2))+1,99)', 'CELL("contents",(R!A3,R!A1))', 'CELL("col",R!C1:D2 R!D2:E3)', 'CELL("format",D!A6)', 'CELL("type",D!A6)']],
  sheet: ['SHEET', ["_xlfn.SHEET('R:Two Words'!A1)=_xlfn.SHEET(R!A1)", '_xlfn.SHEET({1,2})', '_xlfn.SHEET((R!A1,R!B1))=_xlfn.SHEET(R!A1)', "_xlfn.SHEETS('R:Two Words'!A1)", '_xlfn.SHEETS({1})', '_xlfn.SHEETS(TRUE)', '_xlfn.SHEETS((R!A1,R!B2))']],
  formulas: ['FORMULATEXT', ['INDEX(_xlfn.FORMULATEXT(R!A1:C1),1,3)', 'ROWS(_xlfn.FORMULATEXT(R!C1:C3))', 'INDEX(_xlfn.ISFORMULA(R!A1:C1),1,3)', '_xlfn.FORMULATEXT((R!C2,R!C1))', '_xlfn.ISFORMULA((R!A1,R!C1))', "_xlfn.FORMULATEXT('R:Two Words'!C1)", "_xlfn.ISFORMULA('R:Two Words'!C1)", 'SUM(--_xlfn.ISFORMULA(R!A1:E3))']],
  trim: ['TRIMRANGE', ['ROWS(_xlfn.TRIMRANGE(R!G1:H9,2.0))', 'ROWS(_xlfn.TRIMRANGE(R!G1:H9,TRUE))', 'ROWS(_xlfn.TRIMRANGE(R!G1:H9,"2"))', 'ROWS(_xlfn._TRO_ALL(R!A:A))', 'SUM(_xlfn._TRO_TRAILING(R!E1:E9))', 'COLUMNS(_xlfn._TRO_LEADING(R!F1:J1))', 'ROWS(_xlfn._TRO_ALL(R!K1:K9))']],
}

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, [fn, formulas]] of Object.entries(CASES)) add(group, fn, ...formulas)
  return out
}

/** The engine's answers by case id, committed before Excel ran the set (`predict.ts`). */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./indirect-held.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel for Mac did not confirm: a spill reference typed in INDIRECT's text (a known difference), CELL
 * over a reference across several sheets (`#VALUE!`, which the engine answers since), and two TRIMRANGE cases the
 * engine predicted while it took TRIMRANGE's range under ROWS for its place alone (fixed since).
 */
export const MISSED: readonly string[] = ['again#7', 'again#8', 'again#9', 'cell#2', 'trim#1', 'trim#3']
