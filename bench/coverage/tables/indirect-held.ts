/**
 * Held out: INDIRECT, OFFSET and CELL in books the case sets cannot write, predicted by the engine before Excel computed
 * them (`indirect-held.predicted.json`, from `predict.ts`). Defined names (a reference, a range, an OFFSET over a
 * cell, an INDIRECT), a table's columns, cells typed after opening that the found references read or that change where
 * they point, a chain read backwards, rows inserted and deleted and a sheet renamed and inserted (INDIRECT's text stays,
 * OFFSET's reference moves), the looks CELL tells (number formats, alignment, protection, column widths) and the
 * active cell CELL reads without a reference.
 */
import type { CellStyle, StructureOp } from '../../../src/contract.js'
import { book, on, sales, type TableCase, type TableStep, typedFormula } from './cases.js'

const op = (o: StructureOp): TableStep => ({ op: o })
const typed = (cell: string, value: number | string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value } })

/** Number formats by row of column A of sheet F, each over -1234.5. */
const FORMATS = [
  'General', '0', '0.00', '#,##0', '#,##0.00', '$#,##0_);($#,##0)', '$#,##0_);[Red]($#,##0)', '$#,##0.00_);($#,##0.00)', '$#,##0.00_);[Red]($#,##0.00)',
  '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', 'm/d/yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yy h:mm',
  '@', '0.000', '#,##0.0', '(0)', '0;(0)', '[Red]0', '0;[Red]-0', '0;[Blue]0', '"€"#,##0.00', '[$€-407] #,##0.00', 'yyyy-mm-dd', 'mm/dd', 'dd/mm/yyyy', 'mmmm d, yyyy', '[h]:mm:ss',
  '0.0E+0', '#,##0;(#,##0)', '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '0.00;-0.00;"zero"', '000-00-0000',
]

/** Styles: alignments (0 left, 1 right, 2 center, 3 fill, 4 justify), unlocked (5). */
const STYLES: CellStyle[] = [
  { alignment: { horizontal: 'left' } },
  { alignment: { horizontal: 'right' } },
  { alignment: { horizontal: 'center' } },
  { alignment: { horizontal: 'fill' } },
  { alignment: { horizontal: 'justify' } },
  { protection: { locked: false } },
]

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, b: TableCase['book'], steps: TableStep[], read: TableCase['read']) => cases.push({ id, about, book: b, ...(steps.length ? { steps } : {}), read })

  const named = book(['S', { A1: 1, A2: 2, A3: 3, A4: 2, B1: 'A2', C1: '=INDIRECT("rate")', C2: '=SUM(INDIRECT("span"))', C3: '=SUM(INDIRECT("grow"))', C4: '=SUM(grow)', C5: '=pick*10', C6: '=INDIRECT("S!rate")', C7: '=INDIRECT("nosuch")', C8: '=ROWS(grow)', C9: '=INDIRECT("rate")+INDIRECT("Rate")' }])
  named.names = [{ name: 'rate', ref: "'S'!$A$1" }, { name: 'span', ref: "'S'!$A$1:$A$3" }, { name: 'grow', ref: "OFFSET('S'!$A$1,0,0,'S'!$A$4)" }, { name: 'pick', ref: "INDIRECT('S'!$B$1)" }]
  add('names', 'names of a reference, a range, an OFFSET and an INDIRECT, through INDIRECT and directly', named, [], on('S', 'C1:C9'))
  add('names-typed', 'the cells the names read typed: the OFFSET grows, the INDIRECT points elsewhere', named, [typed('A4', 3), typed('B1', 'A3'), typed('A3', 30)], on('S', 'C1:C9'))

  const t = sales()
  add('table', "a table's columns, areas and the whole table through INDIRECT", book(['S', { F1: '=SUM(INDIRECT("Sales[Qty]"))', F2: '=INDIRECT("Sales[[#Headers],[Price]]")', F3: '=ROWS(INDIRECT("Sales"))', F4: '=ROWS(INDIRECT("Sales[#All]"))', F5: '=SUM(INDIRECT("sales[total]"))', F6: '=INDIRECT("Sales[Nope]")', F7: '=SUM(INDIRECT("Sales[[Qty]:[Price]]"))', F8: '=COLUMNS(INDIRECT("Sales[[Qty]:[Price]]"))' }, t]), [typed('B2', 4)], on('S', 'F1:F8'))

  const edits = book(['S', { A1: 1, A2: 2, A3: 2, B1: 'A1', C1: '=INDIRECT(B1)*10', C2: '=SUM(OFFSET(A1,0,0,A3))', C3: '=INDIRECT("A"&A3)', C4: '=CELL("contents",INDIRECT(B1))', C5: '=SUM(A1:INDIRECT("A"&A3))', D1: '=INDIRECT("D2")+1', D2: '=INDIRECT("D3")+1', D3: '=A1*100' }])
  add('typed', 'typed after opening: the cell INDIRECT reads, its text, the size OFFSET takes, a chain read backwards', edits, [typed('A1', 5), typed('B1', 'A2'), typed('A3', 1), typed('A2', 7)], on('S', 'C1:D5'))
  add('typed-formula', 'a formula typed that INDIRECT reads, then one typed over its text', edits, [typedFormula('A2', 'A1*3'), typedFormula('B1', '"A"&A3')], on('S', 'C1:D5'))

  const moved = book(['S', { A1: 1, A2: 2, A3: 3, B1: '=INDIRECT("A2")', B2: '=OFFSET(A1,1,0)', B3: '=SUM(OFFSET(A1,0,0,3))', B4: '=INDIRECT("T!A1")', B5: '=INDIRECT("New!A1")', B6: '=SUM(A1:OFFSET(A1,2,0))' }], ['T', { A1: 'tee' }])
  add('insert-rows', 'a row inserted above: the text stays and reads the cell that came there, OFFSET moves', moved, [op({ op: 'insertRows', sheet: 'S', row: 1, count: 1 })], on('S', 'A1:B8'))
  add('delete-rows', 'the row of the cell INDIRECT and OFFSET read deleted', moved, [op({ op: 'deleteRows', sheet: 'S', row: 2, count: 1 })], on('S', 'A1:B7'))
  add('rename-sheet', 'the sheet INDIRECT names renamed, and a sheet inserted with the name another text names', moved, [op({ op: 'renameSheet', sheet: 'T', name: 'U' }), op({ op: 'insertSheet', name: 'New' }), typed('A1', 9, 'New')], on('S', 'A1:B7'))

  const fmt: Record<string, number | string | object> = {}
  FORMATS.forEach((code, i) => {
    const r = i + 1
    fmt[`A${r}`] = { value: -1234.5, numberFormat: code }
    fmt[`B${r}`] = `=CELL("format",A${r})`
    fmt[`C${r}`] = `=CELL("color",A${r})`
    fmt[`D${r}`] = `=CELL("parentheses",A${r})`
  })
  const looks = book(['F', fmt as never], [
    'L',
    {
      A1: { value: 't', style: 0 }, A2: { value: 't', style: 1 }, A3: { value: 't', style: 2 }, A4: { value: 't', style: 3 }, A5: { value: 't', style: 4 }, A6: 't', A7: 5, A8: { value: 5, style: 1 },
      B1: '=CELL("prefix",A1)', B2: '=CELL("prefix",A2)', B3: '=CELL("prefix",A3)', B4: '=CELL("prefix",A4)', B5: '=CELL("prefix",A5)', B6: '=CELL("prefix",A6)', B7: '=CELL("prefix",A7)', B8: '=CELL("prefix",A8)',
      C1: '=CELL("protect",A1)', C2: { value: 1, style: 5 }, C3: '=CELL("protect",C2)', C4: '=CELL("protect",Z99)',
      D1: '=INDEX(CELL("width",D1),1,1)', D2: '=INDEX(CELL("width",E1),1,1)', D3: '=INDEX(CELL("width",G1),1,1)', D4: '=INDEX(CELL("width",H1),1,1)', D5: '=INDEX(CELL("width",D1),1,2)', D6: '=INDEX(CELL("width",E1),1,2)', D7: '=COLUMNS(CELL("width",D1))',
    },
  ])
  looks.styles = STYLES
  looks.sheets[1].columns = { E: { width: 20.7109375 }, G: { width: 9.5703125 }, H: { width: 0 } }
  add('looks', "CELL's format codes of 43 number formats, the prefix of aligned text, protection and widths", looks, [], [...on('F', `B1:D${FORMATS.length}`), ...on('L', 'B1:D8')])

  const active = book(['S', { A1: 5, B2: 7, C1: '=MID(CELL("address"),FIND("]",CELL("address"))+1,99)', C2: '=CELL("row")', C3: '=CELL("contents")' }], ['T', { D4: 'there' }])
  active.sheets[1].view = { activeCell: 'D4' }
  active.activeSheet = 'T'
  add('active', 'CELL without a reference reads the active cell of the book saved with sheet T active, at D4', active, [], on('S', 'C1:C3'))
  return cases
}
