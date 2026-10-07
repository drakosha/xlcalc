/**
 * Implicit intersection against Excel: formulas of the old kind (no dynamic array metadata), as a file from an
 * older Excel or another program holds them, each in row 2 of its own column beside data in C1:E10 whose rows
 * differ. Excel 365 computes them taking one value of a range where the engine computes a dynamic array, so their
 * values tell where that happens. `npx tsx bench/coverage/run.ts xlsx intersect` has Excel compute and save the
 * book into `test/excel/xlsx/intersect.xlsx`; `test/xlsx-excel.test.ts` checks that the reader reports every
 * formula whose stored value differs from the engine's.
 */
import { formatAddress } from '../../../src/model.js'
import { zip } from '../../../src/xlsx/zip/write.js'

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

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
const NEWER = /\b(XLOOKUP|SEQUENCE|SORT|UNIQUE|FILTER)\(/g
const PREFIX: Record<string, string> = { XLOOKUP: '_xlfn.XLOOKUP', SEQUENCE: '_xlfn.SEQUENCE', SORT: '_xlfn._xlws.SORT', UNIQUE: '_xlfn.UNIQUE', FILTER: '_xlfn._xlws.FILTER' }
/** The cell of the formula `i`: row 2, from column G on. */
export const formulaCell = (i: number) => formatAddress(1, 6 + i)

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** The formulas as a file of the old kind holds them: plain `<f>`, no values, no metadata. */
export function oldFormulaBook(formulas: string[]): Promise<Uint8Array> {
  const cells = formulas.map((f, i) => `<c r="${formulaCell(i)}"><f>${esc(f.replace(NEWER, (_m, n: string) => `${PREFIX[n]}(`))}</f></c>`).join('')
  const rows = Array.from({ length: 10 }, (_, r) => `<row r="${r + 1}">${['C', 'D', 'E'].map((c, k) => `<c r="${c}${r + 1}"><v>${(r + 1) * (k + 1)}</v></c>`).join('')}${r === 1 ? cells : ''}</row>`)
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${TYPE}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${TYPE}.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="one">S!$C$1</definedName><definedName name="span">S!$C$1:$C$3</definedName></definedNames><calcPr fullCalcOnLoad="1"/></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet ${NS}><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join('')}</sheetData></worksheet>`,
  }
  return zip(Object.entries(parts).map(([name, text]) => ({ name, data: enc.encode(text) })))
}
