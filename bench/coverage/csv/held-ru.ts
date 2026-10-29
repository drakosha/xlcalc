/**
 * CSV files in Excel ru-RU held out from the fitting (`files-ru`), each with a byte order mark, as Excel for Mac reads
 * a file without one in Mac Cyrillic: booleans, errors, formulas and money in Russian, dates with month names; and books
 * in ru-RU Excel saves as CSV. Its predictions were committed before Excel ran it.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { CsvCase, CsvCases } from './cases.js'

const book = (cells: SheetInput['cells']): BookInput => ({ locale: 'ru-RU', sheets: [{ name: 'Лист', cells }] })

const bom = (id: string, text: string): CsvCase => ({ id, text, encoding: 'utf-8-bom' })

const CASES: CsvCase[] = [
  bom('logic', 'ИСТИНА;ложь;Истина;#Н/Д;#ДЕЛ/0!;#ИМЯ?;#ЗНАЧ!;#ССЫЛКА!\r\n'),
  bom('formulas', '2;=A1*3;"=СУММ(A1;1)";=СУММ(A1;1);"=ОКРУГЛ(1,5;0)";+A1*2;-A1;@СУММ(1;2);=ЕСЛИ(A1>1;1;0);=SUM(1)\r\n'),
  bom('money', '5 ₽;5₽;-5 ₽;5,50 ₽;1 234 ₽;5 р.;$5;€5;5 €\r\n'),
  bom('dates', '15-мар-2026;15 марта 2026;мар 2026;15.03.2026 18:30;15.3.26;1.2;31.02.2026;18:30:15\r\n'),
  bom('numbers', '1 234,5;-1,5;1,5%;(5);1,5E+3;1.5;"1,5";0,5;,5\r\n'),
  bom('quoted', '"ИСТИНА";"1,5";"15.03.2026";"а;б";"""к"""\r\n'),
  bom('wrap', '"строка 1\nстрока 2";x\r\n'),
  bom('named', 'а;б\r\n'),
  { id: 'w-formats', book: book({ A1: { value: -1234.5, numberFormat: '#,##0.00;[Red]-#,##0.00' }, B1: { value: 46096, numberFormat: 'd mmmm yyyy' }, C1: { value: false }, D1: { formula: 'NA()' }, E1: { value: 0.75, numberFormat: 20 }, F1: { value: 1.5, numberFormat: '0.000' }, G1: { formula: '1/3' }, H1: { value: 'текст' } }) },
  { id: 'w-formulas', book: book({ A1: { value: 2 }, B1: { formula: 'A1>1' }, C1: { formula: 'A1*1.5' }, D1: { formula: '"а"&"б"' } }) },
]

export function buildCsvCases(): CsvCases {
  return { locale: 'ru-RU', cases: CASES }
}
