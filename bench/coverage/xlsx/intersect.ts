/**
 * Implicit intersection against Excel: formulas of the old kind (no dynamic array metadata), as a file from an
 * older Excel or another program holds them, each in row 2 of its own column beside data in C1:E10 whose rows
 * differ. Excel 365 computes them taking one value of a range where a dynamic array formula would give several, so
 * their values tell where that happens. `npx tsx bench/coverage/run.ts xlsx intersect` has Excel compute and save the
 * book into `test/excel/xlsx/intersect/`; `test/xlsx-excel.test.ts` checks that the formulas the reader reads (with
 * `@` where Excel puts it) compute the values Excel stored.
 */
import { formatAddress } from '../../../src/model.js'
import { XLFN, XLWS } from '../../../src/xlsx/formulas/prefixes.js'
import { handBook } from './hand.js'

/** Formulas in row 2 of sheet S from column G on, over data in C1:E10 and the names `span` (C1:C3) and `one` (C1). */
export const OLD_FORMULAS = [
  'C1:C3*2', 'C1:C3', '-C1:C3', 'C1:C3%', 'C1:C3&"x"', 'C1:C3=1', '(C1:C3)', 'C1*2', 'C:C+0', '1:1+0',
  'SUM(C1:C3)', 'SUM(C1:C3*2)', 'SUMPRODUCT(C1:C3*2)', 'ABS(C1:C3)', 'ABS(C1)', 'LEFT(C1:C3,1)', 'DATE(C1:C3,1,1)', 'ROUND(C1:C3,1)',
  'COUNTIF(C1:C3,">0")', 'COUNTIF(C1:C3,D1:D3)', 'SUMIFS(C1:C3,D1:D3,">0")', 'AVERAGE(C1:E3)', 'MAX(C:C)', 'COUNTA(C1:E10)',
  'INDEX(C1:E5,2,1)', 'INDEX(C1:E5,0,1)', 'INDEX(C1:C5,2)', 'VLOOKUP(1,C1:E5,2,FALSE)', 'HLOOKUP(1,C1:E5,2,FALSE)', 'MATCH(1,C1:C5,0)',
  'LOOKUP(1,C1:C5,D1:D5)', 'OFFSET(C1,1,0)', 'OFFSET(C1,0,0,2,1)', 'ROW()', 'ROW(C3)', 'ROW(C1:C3)', 'COLUMN(D1)', 'COLUMN(C1:E1)', 'ROWS(C1:C3)',
  'TRANSPOSE(C1:C3)', 'MMULT(C1:D2,C1:D2)', 'IF(C1>0,1,2)', 'IF(C1:C3>0,1,2)', 'IF(TRUE,C1:C3)', 'SUM(IF(C1:C3>0,1,0))', 'SUM(IF(TRUE,C1:C3))',
  'IFERROR(C1:C3,0)', 'IFERROR(C1,0)', 'CHOOSE(1,C1:C3,D1:D3)', 'N(C1:C3)', 'T(C1:C3)', 'span*1', 'one*1', 'SUM(span)', 'span',
  '{1,2,3}', '{1,2}+0', 'SUM({1,2})', 'ISNUMBER(C1:C3)', 'AND(C1:C3>0)', 'CONCATENATE(C1:C3)', 'TEXT(C1:C3,"0")', 'MEDIAN(C1:C3)',
  'FREQUENCY(C1:C5,D1:D2)', 'EOMONTH(C1:C3,0)', 'NETWORKDAYS(C1,C2:C3)', 'XLOOKUP(1,C1:C5,D1:D5)', 'XLOOKUP(1,C1:C5,D1:E5)', 'SEQUENCE(3)',
  'SORT(C1:C3)', 'UNIQUE(C1:C3)', 'FILTER(C1:C3,C1:C3>1)', 'C1:C3 D2:E2',
]

/** A function newer than Excel 2007 with the prefix a file writes before it (the reader's and writer's table). */
const prefixed = (f: string) => f.replace(/\b([A-Z][A-Z0-9.]*)\(/g, (m, n: string) => (XLWS.has(n) ? `_xlfn._xlws.${n}(` : XLFN.has(n) ? `_xlfn.${n}(` : m))
/** The cell of the formula `i`: row 2, from column G on. */
export const formulaCell = (i: number) => formatAddress(1, 6 + i)

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** The formulas as a file of the old kind holds them: plain `<f>`, no values, no metadata. */
export function oldFormulaBook(formulas: string[]): Promise<Uint8Array> {
  const cells = formulas.map((f, i) => `<c r="${formulaCell(i)}"><f>${esc(prefixed(f))}</f></c>`).join('')
  const rows = Array.from({ length: 10 }, (_, r) => `<row r="${r + 1}">${['C', 'D', 'E'].map((c, k) => `<c r="${c}${r + 1}"><v>${(r + 1) * (k + 1)}</v></c>`).join('')}${r === 1 ? cells : ''}</row>`)
  return handBook({
    sheet: 'S',
    body: `<sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join('')}</sheetData>`,
    workbook: '<definedNames><definedName name="one">S!$C$1</definedName><definedName name="span">S!$C$1:$C$3</definedName></definedNames><calcPr fullCalcOnLoad="1"/>',
  })
}
