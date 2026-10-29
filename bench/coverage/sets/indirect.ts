/**
 * INDIRECT and OFFSET in every form, and the functions about references and the book: A1 and R1C1 text, another sheet,
 * a quoted sheet, whole columns and rows, a spill reference, a text that names nothing, OFFSET with a size, a negative
 * size and off the sheet's edges, both as arguments of SUM, INDEX, ROWS and the range operator, lifted over arrays;
 * CELL's answers that need no look, INFO's constant ones, SHEET, SHEETS, FORMULATEXT and ISFORMULA; TRIMRANGE, and
 * the trim references as a file stores them (`_xlfn._TRO_ALL` and the others). The case sheet C comes after
 * the set's sheets and D in Excel's book and after D and them in the engine's, so a sheet's number is read against C's.
 * Answers: `test/excel/cases/indirect.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** R: numbers, texts that name cells, a formula, an array spilled from E1; Two Words: a sheet whose name needs quotes. */
export const SHEETS: SheetCells = {
  R: {
    A1: 10, A2: 20, A3: 30, A4: 'x', A6: 60,
    B1: 'R!A2', B2: 'A3', B3: 'R!R1C1', B4: "'Two Words'!A1", B5: 'R!E1#', B6: 'nowhere!A1', B7: 'R!A1:A3',
    C1: { formula: 'A1*2' }, C2: { formula: 'SUM(A1:A3)' }, C3: { formula: '"t"&A4' },
    E1: { formula: '_xlfn.SEQUENCE(3)', dynamic: true },
    G2: 5, G3: '', H4: 7, H5: { formula: '""' },
  },
  'Two Words': { A1: 7, B2: 8 },
}

const CASES: Record<string, [fn: string, formulas: string[]]> = {
  a1: ['INDIRECT', ['INDIRECT("R!A1")', 'INDIRECT(R!B1)', 'INDIRECT("r!a1")', 'INDIRECT("R!$A$3")', 'INDIRECT(R!B4)', "INDIRECT(\"'Two Words'!B2\")", 'INDIRECT("Two Words!A1")', 'SUM(INDIRECT(R!B7))', 'ROWS(INDIRECT("R!A:A"))', 'COLUMNS(INDIRECT("R!1:1"))', 'SUM(INDIRECT("R!A:A"))', 'INDIRECT("R!C1")', 'INDIRECT("R!C3")', 'INDIRECT("R!A5")', 'INDIRECT(R!B6)', 'INDIRECT("")', 'INDIRECT(" R!A1")', 'INDIRECT("R!A1 ")', 'INDIRECT("=R!A1")', 'INDIRECT("R!A1+1")', 'INDIRECT("(R!A1)")', 'INDIRECT("R!A1,R!A2")', 'INDIRECT("R!A1:R!A2")', 'INDIRECT(1)', 'INDIRECT(TRUE)', 'INDIRECT(D!A7)', 'INDIRECT("R!XFD1048576")', 'INDIRECT("R!XFE1")', 'INDIRECT("R!A1048577")', 'INDIRECT("R:A1")', 'INDIRECT("A1",TRUE)', 'SUM(INDIRECT("R!E1#"))', 'ROWS(INDIRECT(R!B5))', 'INDIRECT("R!A1#")', 'INDIRECT("R!A1",FALSE)', 'INDIRECT("[1]R!A1")', 'INDIRECT("R:R!A1")', 'INDIRECT("R!A1","x")', 'ISREF(INDIRECT("R!A1"))', 'ISREF(INDIRECT("nowhere"))']],
  r1c1: ['INDIRECT', ['INDIRECT(R!B3,FALSE)', 'INDIRECT("R!R2C1",FALSE)', 'INDIRECT("r!r3c1",0)', 'SUM(INDIRECT("R!R1C1:R3C1",FALSE))', 'ROWS(INDIRECT("R!R1:R3",FALSE))', 'COLUMNS(INDIRECT("R!C1:C3",FALSE))', 'ROW(INDIRECT("R",FALSE))', 'COLUMN(INDIRECT("C",FALSE))', 'ROW(INDIRECT("R[2]C[3]",FALSE))', 'COLUMN(INDIRECT("RC[3]",FALSE))', 'ROW(INDIRECT("R[-1]C",FALSE))', 'INDIRECT("R!R0C1",FALSE)', 'INDIRECT("R!R1C16385",FALSE)', 'INDIRECT("R!A1",FALSE)', 'INDIRECT("R!R1C1:R2",FALSE)', "INDIRECT(\"'Two Words'!R2C2\",FALSE)", 'INDIRECT("R!R1C[1]",FALSE)', 'INDIRECT("R!R[1]C1",FALSE)']],
  offset: ['OFFSET', ['OFFSET(R!A1,1,0)', 'OFFSET(R!A1,0,2)', 'SUM(OFFSET(R!A1,0,0,3,1))', 'SUM(OFFSET(R!A1,0,0,3))', 'ROWS(OFFSET(R!A1:A3,1,0))', 'COLUMNS(OFFSET(R!A1,0,0,1,5))', 'OFFSET(R!A1,-1,0)', 'OFFSET(R!A1,0,-1)', 'OFFSET(R!A1,0,0,0)', 'OFFSET(R!A1,0,0,1,0)', 'SUM(OFFSET(R!A3,0,0,-2))', 'ROWS(OFFSET(R!A3,0,0,-2))', 'SUM(OFFSET(R!C1,0,0,1,-3))', 'OFFSET(R!A1,1.9,0)', 'OFFSET(R!A3,-1.9,0)', 'OFFSET(R!A1,0.99999999,0)', 'SUM(OFFSET(R!A1,0,0,2.7))', 'OFFSET(R!A1,"1",0)', 'OFFSET(R!A1,"x",0)', 'OFFSET(R!A1,TRUE,0)', 'OFFSET(R!A1,,)', 'SUM(OFFSET(R!A1,0,0,,))', 'ROWS(OFFSET(R!A1:B2,1,1))', 'OFFSET((R!A1,R!A2),0,0)', 'OFFSET(D!A7,0,0)', 'OFFSET(R!A1,1048575,0)', 'OFFSET(R!A1,1048576,0)', 'OFFSET(R!A1,0,16383)', 'OFFSET(R!A1,0,0,1048577)', 'ROWS(OFFSET(R!A1,0,0,1048576))', 'OFFSET(R!A1,R!H4-6,0)', 'SUM(OFFSET(R!A1,{0,1,2},0))', 'SUM(OFFSET(R!A1,{0;1},0,2))', 'INDEX(OFFSET(R!A1,0,0,3),2)', 'SUM(R!A1:OFFSET(R!A1,2,0))', 'ROWS(R!A1:OFFSET(R!A1,9,0))', 'OFFSET(INDIRECT("R!A1"),2,0)', 'ISREF(OFFSET(R!A1,0,0))', 'AREAS(OFFSET(R!A1,0,0))', 'OFFSET(R!E1#,1,0)', 'SUM(OFFSET(R!E1#,1,0))']],
  arrays: ['INDIRECT', ['SUM(INDIRECT({"R!A1","R!A2"}))', 'SUM(INDIRECT({"R!A1:A2","R!A3"}))', 'SUMPRODUCT(ROW(INDIRECT("1:4")))', 'SUM(N(INDIRECT({"R!A1","R!A3"})))', 'ROWS(INDIRECT("R!A1:A3"))*COLUMNS(INDIRECT("R!A1:C1"))']],
  cell: ['CELL', ['CELL("address",R!B3)', 'CELL("address",C!B3)', 'CELL("address",R!A1:C3)', 'MID(CELL("address",R!C2),FIND("]",CELL("address",R!C2))+1,99)', 'CELL("row",R!B7)', 'CELL("col",R!H4)', 'CELL("contents",R!A2)', 'CELL("contents",R!C1)', 'CELL("contents",R!A5)', 'CELL("contents",R!A1:A3)', 'CELL("type",R!A1)', 'CELL("type",R!A4)', 'CELL("type",R!A5)', 'CELL("type",R!C3)', 'CELL("type",R!H5)', 'CELL("type",D!A7)', 'CELL("TYPE",R!A1)', 'CELL(" type",R!A1)', 'CELL("bogus",R!A1)', 'CELL(1,R!A1)', 'CELL("row",(R!A1,R!B2))', 'CELL("row",R:R!A1)', 'CELL("row",R!C1:D2 R!D2:E3)', 'CELL("contents",D!A7)', 'ISTEXT(CELL("filename",R!A1))', 'CELL("row",INDIRECT("R!A5"))', 'CELL("contents",OFFSET(R!A1,2,0))']],
  info: ['INFO', ['INFO("system")', 'INFO("release")', 'INFO("recalc")', 'ISTEXT(INFO("osversion"))', 'ISTEXT(INFO("directory"))', 'ISNUMBER(INFO("numfile"))', 'INFO("origin")', 'INFO("memavail")', 'INFO("totmem")', 'INFO("SYSTEM")', 'INFO("bogus")', 'INFO(1)']],
  sheet: ['SHEET', ['SHEET()', 'SHEET(C!A1)', 'SHEET(R!A1)<SHEET()', 'SHEET("C")', 'SHEET("c")', 'SHEET("R")=SHEET(R!A1)', 'SHEET("nowhere")', 'SHEET(1)', 'SHEET(TRUE)', 'SHEET(D!A7)<SHEET()', 'SHEET(R:C!A1)=SHEET(R!A1)', 'SHEET((R!A1,R!A2))=SHEET(R!A1)', 'SHEET(INDIRECT("R!A1"))=SHEET(R!A1)', "SHEET(\"'Two Words'\")", 'SHEET("Two Words")=SHEET(\'Two Words\'!A1)', 'SHEET({"C","R"})']],
  sheets: ['SHEETS', ['SHEETS()', 'SHEETS(R!A1)', "SHEETS(R:'Two Words'!A1)", 'SHEETS((R!A1,R!A2))', 'SHEETS("R")', 'SHEETS(1)', 'SHEETS(D!A7)']],
  formulatext: ['FORMULATEXT', ['FORMULATEXT(R!C1)', 'FORMULATEXT(R!C2)', 'FORMULATEXT(R!C3)', 'FORMULATEXT(R!A1)', 'FORMULATEXT(R!A5)', 'FORMULATEXT(R!E1)', 'FORMULATEXT(R!E2)', 'FORMULATEXT(R!C1:C3)', 'FORMULATEXT(R!H5)', 'FORMULATEXT(INDEX(R!C1:C3,2))', 'FORMULATEXT(IF(TRUE,R!C3))', 'FORMULATEXT((R!C1,R!C2))', 'FORMULATEXT(R:C!C1)', 'FORMULATEXT(INDIRECT("R!C2"))']],
  isformula: ['ISFORMULA', ['ISFORMULA(R!C1)', 'ISFORMULA(R!A1)', 'ISFORMULA(R!A5)', 'ISFORMULA(R!E1)', 'ISFORMULA(R!E2)', 'ISFORMULA(R!C1:C3)', 'ISFORMULA(R!A1:C1)', 'ISFORMULA(INDEX(R!A1:C1,3))', 'ISFORMULA(CHOOSE(2,R!A1,R!C1))', 'ISFORMULA((R!C1,R!A1))', 'ISFORMULA(R:C!C1)', 'ISFORMULA(D!A7)']],
  trimrange: ['TRIMRANGE', ['ROWS(TRIMRANGE(R!A1:A9))', 'SUM(TRIMRANGE(R!A1:A9))', 'ROWS(TRIMRANGE(R!A:A))', 'ROWS(TRIMRANGE(R!G1:H9))', 'COLUMNS(TRIMRANGE(R!F1:J9))', 'ROW(TRIMRANGE(R!G1:H9))', 'COLUMN(TRIMRANGE(R!F1:J9))', 'ROWS(TRIMRANGE(R!G1:H9,1))', 'ROWS(TRIMRANGE(R!G1:H9,2))', 'ROWS(TRIMRANGE(R!G1:H9,0))', 'COLUMNS(TRIMRANGE(R!F1:J9,3,1))', 'COLUMNS(TRIMRANGE(R!F1:J9,3,2))', 'COLUMNS(TRIMRANGE(R!F1:J9,3,0))', 'TRIMRANGE(R!K1:K9)', 'TRIMRANGE(R!G1:H9,4)', 'TRIMRANGE(R!G1:H9,-1)', 'TRIMRANGE(R!G1:H9,"x")', 'ROWS(TRIMRANGE(R!H4:H9))', 'TRIMRANGE({1,2})', 'TRIMRANGE(5)', 'TRIMRANGE((R!A1,R!A2))', 'ROWS(TRIMRANGE(R!G1:H9,1.5))', 'SUM(TRIMRANGE(R!E1:E9))']],
  tro: ['TRIMRANGE', ['ROWS(_xlfn._TRO_ALL(R!G1:H9))', 'ROWS(_xlfn._TRO_LEADING(R!G1:H9))', 'ROWS(_xlfn._TRO_TRAILING(R!G1:H9))', 'ROWS(_xlfn._TRO_TRAILING(R!A:A))']],
}

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, [fn, formulas]] of Object.entries(CASES)) add(group, fn, ...formulas)
  return out
}

/** The engine's answers by case id, committed before Excel ran the set (`predict.ts`). */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./indirect.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel for Mac did not confirm, which the fit took up where both Excels agree: OFFSET truncating its
 * numbers (`0.99999999` is 0) and taking a size left out as the reference's; CELL, SHEET, FORMULATEXT and ISFORMULA over
 * a union's first area; FORMULATEXT and ISFORMULA over a range cell by cell; SHEET and SHEETS of a value `#N/A`; TRIMRANGE's
 * codes whole numbers; the trim references' file form, which Excel for Mac computes; a spill reference from a cell under
 * ROWS ordered after its array. Known differences: a spill reference
 * in INDIRECT's typed text, CELL's book name, INFO's constants, a quoted last sheet; `#NUM!` the Mac did not repeat
 * (`test/excel-open.ts`).
 */
export const MISSED: readonly string[] = ['a1#26', 'a1#32', 'a1#33', 'offset#16', 'offset#22', 'offset#37', 'arrays#1', 'arrays#2', 'arrays#3', 'arrays#4', 'arrays#5', 'cell#1', 'cell#3', 'cell#4', 'cell#21', 'cell#22', 'cell#26', 'info#1', 'info#2', 'info#10', 'sheet#11', 'sheet#12', 'sheet#16', 'sheets#3', 'sheets#5', 'sheets#6', 'formulatext#4', 'formulatext#8', 'formulatext#12', 'formulatext#13', 'isformula#6', 'isformula#7', 'isformula#10', 'trimrange#22', 'tro#1', 'tro#2', 'tro#3']
