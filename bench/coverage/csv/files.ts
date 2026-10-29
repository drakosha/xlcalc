/**
 * CSV files in Excel en-US, the set the reader was fitted to: fields and quotes, line breaks, empty fields and rows,
 * numbers, dates, times, percentages and money as Excel reads them, booleans and errors, formulas, leading zeros and long
 * numbers, an apostrophe, spaces, `sep=` with and without a byte order mark, other
 * delimiters, encodings, the sheet's name from the file's; and books Excel saves as CSV: formats, a narrow column, the
 * used range, short rows, sheets, an array, characters the code page lacks.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { CsvCase, CsvCases } from './cases.js'

const book = (cells: SheetInput['cells'], more: Partial<SheetInput> = {}): BookInput => ({ sheets: [{ name: 'S', cells, ...more }] })

const CASES: CsvCase[] = [
  { id: 'basic', text: 'Item,Qty,Price\r\nApple,3,1.5\r\nPear,10,0.25\r\n' },
  { id: 'quotes', text: '"a,b","say ""hi""","line1\nline2",plain"quote,"x"y,"open\r\nnext,row\r\n' },
  { id: 'quote-unclosed', text: 'a,"b\r\nc,d\r\ne\r\n' },
  { id: 'quote-spaced', text: 'a, "b,c" ,d\r\n"",""""\r\n' },
  { id: 'breaks', text: 'a,b\nc,d\re,f\r\ng' },
  { id: 'empty', text: 'a,,c\r\n\r\n,,\r\nd\r\n\r\n' },
  { id: 'numbers', text: '1234.5,"1,234.5",1e3,-5,(5),5%,$1.50,1/2,0 1/2, 12 ,+5,.5,-.5,1.,"1,2,3"\r\n' },
  { id: 'numbers-more', text: '1E+3,1.5e-3,12345678901,1e308,1e309,"1 234",50%,5.5%,-5%,$-5,-$5,"$1,234",€5,£5\r\n' },
  { id: 'dates', text: '3/15/2026,2026-03-15,15-Mar-2026,Mar 15,3/15,18:30,6:30 PM,3/15/2026 18:30,25:00,1:30:45,1/4,2026/3/15\r\n' },
  { id: 'logic', text: 'TRUE,false,True,#N/A,#DIV/0!,#ref!,#NAME?,#VALUE\r\n' },
  { id: 'formulas', text: '=1+2,=SUM(1,2),"=SUM(1,2)",+1+2,-1+2,@SUM(1),=1+,=,=A1*2,=foo(1),-A1,+A1\r\n' },
  { id: 'codes', text: '00123,007,0,0.5,0123.5,1234567890123456789,123456789012345,1234567890123456,+0012,-007\r\n' },
  { id: 'apostrophe', text: "'007,'abc,'=1+2,a'b\r\n" },
  { id: 'spaces', text: ' abc ,  , 5 ,"  7  ",x \r\n' },
  { id: 'sep', text: 'sep=;\r\na;b,c\r\n1;2,5\r\n' },
  { id: 'sep-bom', text: 'sep=;\r\na;b,c\r\n1;2,5\r\n', encoding: 'utf-8-bom' },
  { id: 'sep-tab', text: 'sep=\t\r\na\tb,c\r\n' },
  { id: 'semicolons', text: 'a;b;c\r\n1;2;3\r\n' },
  { id: 'tabs', text: 'a\tb\tc\r\n1\t2\t3\r\n' },
  { id: 'utf8', text: 'Café,Привет,€,日本\r\n' },
  { id: 'utf8-bom', text: 'Café,Привет,€,日本\r\n', encoding: 'utf-8-bom' },
  { id: 'utf16le-bom', text: 'a,b\tc\r\nCafé,Привет\r\n', encoding: 'utf-16le-bom' },
  { id: 'utf16be-bom', text: 'a,b\tc\r\nCafé,Привет\r\n', encoding: 'utf-16be-bom' },
  { id: 'utf16le', text: 'a,b\tc\r\nCafé,Привет\r\n', encoding: 'utf-16le' },
  { id: 'cp1252', text: 'Café,naïve,€5,“q”\r\n', encoding: 'windows-1252' },
  { id: 'cp1251', text: 'Привет,мир\r\n', encoding: 'windows-1251' },
  { id: 'named', text: 'a,b\r\n', name: 'Sales 2026.csv' },
  { id: 'named-long', text: 'a,b\r\n', name: 'A very long file name of more than 31 characters.csv' },
  { id: 'empty-file', text: '' },
  // books Excel saves as CSV
  {
    id: 'w-formats',
    book: book({
      A1: { value: 1234.5, numberFormat: '#,##0.00' },
      B1: { value: 46096, numberFormat: 14 },
      C1: { value: 46096, numberFormat: 'dddd, mmmm d, yyyy' },
      D1: { value: 1.5, numberFormat: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' },
      E1: { value: 0.125, numberFormat: '0.0%' },
      F1: { value: 1.23456789012345 },
      G1: { value: 123456789012345680000 },
      H1: { value: true },
      I1: { formula: '1/0' },
      J1: { value: 0.770833333333333, numberFormat: 'h:mm' },
      K1: { value: -5, numberFormat: '0;(0)' },
      L1: { value: 1234567.891 },
    }),
  },
  { id: 'w-texts', book: book({ A1: { value: 'a,b' }, B1: { value: 'say "hi"' }, C1: { value: 'line1\nline2' }, D1: { value: ' lead' }, E1: { value: 'trail ' }, F1: { value: '=1+2' }, G1: { value: '007' }, H1: { value: "'q" }, I1: { value: 'a;b' } }) },
  { id: 'w-wide', book: book({ A1: { value: 46096, numberFormat: 'dddd, mmmm d, yyyy' }, B1: { value: 1.23456789012345 } }, { columns: { A: { width: 30 }, B: { width: 20 } } }) },
  { id: 'w-offset', book: book({ B2: { value: 'b2' }, C3: { value: 3 } }) },
  { id: 'w-short-rows', book: book({ ...Object.fromEntries(Array.from({ length: 20 }, (_, i) => [`A${i + 1}`, { value: i + 1 }])), C1: { value: 'c' } }) },
  { id: 'w-gap', book: book({ A1: { value: 1 }, A3: { value: 3 }, C2: { value: 'c' } }) },
  { id: 'w-formatted', book: book({ A1: { value: 1 }, D5: { numberFormat: '0.00' } }) },
  { id: 'w-sheets', book: { activeSheet: 'Two', sheets: [{ name: 'One', cells: { A1: { value: 1 } } }, { name: 'Two', cells: { A1: { value: 2 }, B1: { formula: 'One!A1+A1' } } }] } },
  { id: 'w-array', book: book({ A1: { formula: 'SEQUENCE(3,2)' }, D1: { formula: 'A3*10' } }) },
  { id: 'w-unicode', book: book({ A1: { value: 'Привет' }, B1: { value: '€5' }, C1: { value: 'Café' }, D1: { value: '日本' } }) },
  { id: 'w-empty', book: book({}) },
]

export function buildCsvCases(): CsvCases {
  return { cases: CASES }
}
