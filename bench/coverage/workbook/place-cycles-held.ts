/**
 * Held out: a formula over its own cell where the function reads the cell's place alone or its look, not its value
 * (ROW, COLUMNS, ISREF, AREAS, SHEET, CELL's address, row and width, the same through INDIRECT and OFFSET), against
 * one that reads it (CELL's contents, SUM, FORMULATEXT, ISFORMULA, INDIRECT alone): which of them Excel takes as a
 * circular reference. A book a case, its one formula in A1 of S, read after opening and after a full calculation;
 * Excel for Windows alone, as the Mac's step refuses a circular reference. The answers predicted by the engine were
 * committed before Excel ran (`place-cycles-held.predicted.json`).
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { WorkbookCase } from './cases.js'
import { PLAIN, sheet } from './iterate.js'

const FORMULAS: Record<string, string> = {
  row: '=ROW(A1)',
  columns: '=COLUMNS(A1:C1)',
  isref: '=ISREF(A1)',
  areas: '=AREAS(A1)',
  sheet: '=SHEET(A1)',
  'cell-address': '=CELL("address",A1)',
  'cell-row': '=CELL("row",A1)',
  'cell-width': '=INDEX(CELL("width",A1),1,1)',
  'cell-contents': '=CELL("contents",A1)',
  'cell-type': '=CELL("type",A1)',
  sum: '=SUM(A1)+1',
  formulatext: '=FORMULATEXT(A1)',
  isformula: '=ISFORMULA(A1)',
  'row-indirect': '=ROW(INDIRECT("A1"))',
  'rows-offset': '=ROWS(A1:OFFSET(A1,2,0))',
  'offset-below': '=OFFSET(A1,1,0)+1',
  indirect: '=INDIRECT("A1")+1',
}

const PREDICTED = JSON.parse(readFileSync(new URL('./place-cycles-held.predicted.json', import.meta.url), 'utf8')) as Record<string, Record<string, Val>[]>

export function buildWorkbookCases(): WorkbookCase[] {
  return Object.entries(FORMULAS).map(([id, f]): WorkbookCase => {
    const c: WorkbookCase = { id, about: `${f} in A1`, book: { sheets: [sheet('S', { A1: f, A2: 5 })] }, read: [{ sheet: 'S', range: 'A1' }], file: PLAIN, steps: ['read', 'full', 'read'], cycle: true }
    return PREDICTED[id] ? { ...c, predicted: PREDICTED[id] } : c
  })
}
