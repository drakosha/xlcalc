/**
 * CSV files in Excel ru-RU, the set the reader was fitted to: the list separator `;`, numbers with a decimal comma and
 * a space between thousands, dates and times, money, booleans and errors in Russian, formulas in Russian, a comma file,
 * `sep=`, Cyrillic in UTF-8 with and without a byte order mark and in Windows-1251; and books in ru-RU Excel saves as CSV.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { CsvCase, CsvCases } from './cases.js'

const book = (cells: SheetInput['cells']): BookInput => ({ locale: 'ru-RU', sheets: [{ name: 'Лист', cells }] })

const CASES: CsvCase[] = [
  { id: 'basic', text: 'Товар;Кол-во;Цена\r\nЯблоко;3;1,5\r\nГруша;10;0,25\r\n' },
  { id: 'numbers', text: '1234,5;1 234,5;"1 234,5";1.5;1e3;-5;(5);5%;5,5%;1/2;0 1/2; 12 ;+5;,5\r\n' },
  { id: 'money', text: '5 ₽;5₽;5 р.;5р.;$5;€5;-5 ₽\r\n' },
  { id: 'dates', text: '15.03.2026;15.03.26;15.03;2026-03-15;15/03/2026;15-мар-2026;18:30;6:30 PM;15.03.2026 18:30;25:00;1:30:45\r\n' },
  { id: 'logic', text: 'ИСТИНА;ложь;TRUE;#Н/Д;#ДЕЛ/0!;#N/A;#ИМЯ?\r\n' },
  { id: 'formulas', text: '=1+2;"=СУММ(1;2)";=СУММ(1;2);=SUM(1;2);"=ОКРУГЛ(1,5;0)";+1+2;-A1;=1,5*2\r\n' },
  { id: 'commas', text: 'a,b,c\r\n1,5,2\r\n' },
  { id: 'sep', text: 'sep=,\r\na,b;c\r\n1,2;5\r\n' },
  { id: 'tabs', text: 'a\tb\r\n1\t2\r\n' },
  { id: 'utf8', text: 'Привет;мир;Café\r\n' },
  { id: 'utf8-bom', text: 'Привет;мир;Café\r\n', encoding: 'utf-8-bom' },
  { id: 'cp1251', text: 'Привет;мир;№5\r\n', encoding: 'windows-1251' },
  { id: 'codes', text: '00123;007;1234567890123456789\r\n' },
  { id: 'named', text: 'a;b\r\n', name: 'Отчёт за март.csv' },
  // books Excel saves as CSV
  {
    id: 'w-formats',
    book: book({
      A1: { value: 1234.5, numberFormat: 4 },
      B1: { value: 46096, numberFormat: 14 },
      C1: { value: 1.5 },
      D1: { value: true },
      E1: { formula: '1/0' },
      F1: { value: 0.125, numberFormat: '0.0%' },
      G1: { value: 'a;b' },
      H1: { value: 'a,b' },
      I1: { value: 1.5, numberFormat: '#,##0.00 [$₽-419]' },
    }),
  },
  { id: 'w-unicode', book: book({ A1: { value: 'Привет' }, B1: { value: '€5' }, C1: { value: 'Café' }, D1: { value: '№' } }) },
]

export function buildCsvCases(): CsvCases {
  return { locale: 'ru-RU', cases: CASES }
}
