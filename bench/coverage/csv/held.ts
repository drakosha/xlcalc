/**
 * CSV files in Excel en-US held out from the fitting (`files`): its predictions were committed before Excel ran it.
 * Fields begun by `@`, more quotes, numbers, dates and times, the euro in a file with a byte order mark, `sep=` in other
 * forms, a wrapped field at a CR; books with more formats, a fill, a fraction, hidden columns, a cell formatted before
 * the data.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { CsvCase, CsvCases } from './cases.js'

const book = (cells: SheetInput['cells'], more: Partial<SheetInput> = {}): BookInput => ({ sheets: [{ name: 'S', cells, ...more }] })

const CASES: CsvCase[] = [
  { id: 'at', text: '1,2\r\n@A1,@abc,@1,@,@A1+B1,@SUM(A1:B1)\r\n' },
  { id: 'quoted-values', text: '"00123","1,234","TRUE","=1+2"," 5","5 "\r\n' },
  { id: 'numbers', text: '1.5E+3,1 1/2,(5%),-0,1.0,"12,345,678",1e-400,0x1F,1_000,5-\r\n' },
  { id: 'dates', text: 'Mar 15 2026,15 March 2026,1-2,2026-13-01,12:00 AM,0:00,23:59:59,3/15/26,15-Mar,"March 15, 2026"\r\n' },
  { id: 'words', text: 'TRUE ,FALSE,yes,#N/A ,#NULL!,#SPILL!,#CALC!,#GETTING_DATA\r\n' },
  { id: 'euro-bom', text: '€5,5 €,€-5,-€5,£5,¥5\r\n', encoding: 'utf-8-bom' },
  { id: 'sep-upper', text: 'SEP=;\r\na;b\r\n' },
  { id: 'sep-comma', text: 'sep=,\r\na,b\r\n' },
  { id: 'sep-late', text: 'a,b\r\nsep=;\r\n' },
  { id: 'cr-wrap', text: '"a\rb",c\r\n' },
  { id: 'formulas', text: '2,=A1*3,"=""a""&""b""",+A1*2,=IF(A1>1,""big"",""small""),"=IF(A1>1,""big"",""small"")",=A1:A2\r\n' },
  { id: 'ragged', text: 'a\r\nb,c,d\r\ne,f\r\n' },
  { id: 'w-formats', book: book({ A1: { value: 1.5, numberFormat: '0.000' }, B1: { value: 1.25, numberFormat: '[h]:mm' }, C1: { value: 12, numberFormat: '@' }, D1: { value: 1234, numberFormat: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)' }, E1: { value: 1.5, numberFormat: '# ?/?' }, F1: { value: -1234.5, numberFormat: '#,##0.00;[Red]-#,##0.00' }, G1: { value: 0, numberFormat: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)' }, H1: { value: 'x', numberFormat: '"<"@">"' } }) },
  { id: 'w-formulas', book: book({ A1: { value: 2 }, B1: { formula: 'A1>1' }, C1: { formula: 'NA()' }, D1: { formula: '"a"&"b"' }, E1: { formula: 'A1/3' }, F1: { formula: '""' } }) },
  { id: 'w-hidden', book: book({ A1: { value: 1 }, B1: { value: 2 }, C1: { value: 3 } }, { columns: { B: { hidden: true } } }) },
  { id: 'w-formatted-first', book: book({ A1: { numberFormat: '0.00' }, C3: { value: 'c' } }) },
  { id: 'w-long-text', book: book({ A1: { value: 'x'.repeat(300) }, B1: { value: 'tab\there' } }) },
]

export function buildCsvCases(): CsvCases {
  return { cases: CASES }
}
