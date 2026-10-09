/**
 * Excel cases for computing by locale (`docs/locale.md`, "Data to measure"): text read as a number, a date, a
 * time or a boolean, numbers and booleans turned into text, TEXT with Russian and English format codes,
 * DOLLAR and FIXED, NUMBERVALUE, complex text, criteria text. Computed in Excel ru-RU
 * (`run.ts probes locale ru`, `test/excel/cases/ru-RU/locale.json`), checked by `test/excel-locales-sets.test.ts`.
 * The sheet `L` holds the values the criteria count. Sizes are never taken from text. Computed as dynamic
 * array formulas (Excel 365).
 */
import type { Probe, SheetCells } from '../probes.js'

/** L!A1:A16: numbers, dates (46143 is 2026-05-01, 46023 2026-01-01, 45657 2024-12-31), booleans, text that reads as numbers and words, #N/A. */
export const SHEETS: SheetCells = {
  L: {
    A1: 1,
    A2: 1.5,
    A3: 2,
    A4: 46143,
    A5: 46023,
    A6: 45657,
    A7: true,
    A8: false,
    A9: '1,5',
    A10: '1.5',
    A11: 'ИСТИНА',
    A12: 'TRUE',
    A13: 'abc',
    A14: '1 234,5',
    A15: 1234.5,
    A16: { formula: 'NA()' },
  },
  // one value per cell, each read alone by a criterion
  M: { A1: 'TRUE', B1: true, C1: 'ИСТИНА', D1: '2', E1: 2, F1: '1 234,5', G1: '1.5', H1: '31.01.2025', I1: '50%', J1: '1,5%' },
}

const q = (s: string) => `"${s.replace(/"/g, '""')}"`
const NBSP = ' '

/** Text read as a number: VALUE and an operator (`1.5` and `0.5` are among the dates). */
const NUMBERS = [
  '1,5', '0,5', '-1,5', '+1,5', '1,5%', '%1,5', '(1,5)', '-(1,5)', '1,5E+3', '1,5e3', '1,5E-3', '1E3', ',5', '1,', '1,5,5',
  '1 234,5', `1${NBSP}234,5`, '1 234 567,5', '12 34', '1 2345', '1 234', '1.234,5', '1,234.5', '1234,5 ₽', '1234,5₽', '₽1234,5',
  '-1234,5 ₽', '1 234,5 р.', '1234,5 руб.', '$1,000', '$1000', '1 1/2', '3 1/4', ' 1,5 ', '1,5 %', '0,0001', '123456789012345678',
  '1,99999999999999999', 'ИСТИНА', 'TRUE',
]

/** Date and time text: VALUE, an operator, DATEVALUE, TIMEVALUE and a date argument (DAY). */
const DATES = [
  '31.01.2025', '31.1.2025', '31.01.25', '1.5.29', '1.5.30', '31/01/2025', '1/2/2025', '31-01-2025', '1-2-2025', '2025-01-31',
  '2025/01/31', '2025.01.31', '31 января 2025', '31 январь 2025', '31 янв 2025', '31.янв.2025', '31-янв-2025', '31-янв-25', 'янв 2025',
  'январь 2025', 'янв-25', '15 янв', '1.5', '01.05', '31.12', '29.2', '29.02.2024', '29.02.2025', '32.1', '1.13', '0.5', '1.0', '12.5',
  '15.01', '1/2', '31 Jan 2025', '31-Jan-2025', 'Jan 31, 2025', '1/31/2025', '31 ЯНВАРЯ 2025', '31 Января 2025', '31.01.2025 18:30',
  '31.01.2025 18:30:15', '31.01.2025 6:30 PM', '1.5 18:30', '31.01.2025 25:00', '18:30', '18:30:15', '18:30:15,5', '18:30:15.5',
  '6:30 PM', '6:30 AM', '6 PM', '25:00', '1:61', '6:30 ПП', '6:30 ДП', '0:00',
]

/** More number text: currency, parentheses, signs and groups together. */
const NUMBERS_MORE = [
  '(1 234,5)', '1 234,5 ₽', '-1 234,5 ₽', '(1234,5 ₽)', '1234,5 ₽%', '₽ 1234,5', '1,5₽ ', '- 1,5', '1,5-', '1 1/2 ₽', '1 234,5%', '%(1,5)', '1 234 5', '1  234',
  '1 234,5E+2', '+(1,5)', '1 234,',
]

/** More date text: month and year, two-digit years, separators. */
const DATES_MORE = [
  '1.2025', '5.25', '12.99', '13.13', '1.29', '1.30', '1.99', '1.100', '0.13', '1/13', '1-13', '31-янв', 'янв-31', 'янв.25', '2025-1-31', '1.5.2025 6 PM',
  '31.01.2025 18.30', '31,01,2025', '31 01 2025', '1.5.', '31.01.2025.', '1.5 6 PM', '31.01.2025 18:30 PM',
]

/** Month names in date text: nominative, genitive, short, as `15 <name> 2025` and `15-<name>-2025`. */
const MONTHS = [
  ['январь', 'января', 'янв'], ['февраль', 'февраля', 'фев'], ['март', 'марта', 'мар'], ['апрель', 'апреля', 'апр'], ['май', 'мая', 'май'],
  ['июнь', 'июня', 'июн'], ['июль', 'июля', 'июл'], ['август', 'августа', 'авг'], ['сентябрь', 'сентября', 'сен'], ['октябрь', 'октября', 'окт'],
  ['ноябрь', 'ноября', 'ноя'], ['декабрь', 'декабря', 'дек'],
]
const MONTH_EXTRA = ['февр', 'сент', 'сентяб', 'ноябр', 'Янв', 'ЯНВ', 'jan', 'mar', 'янва', 'январ', 'ма', 'ию', 'июн.', 'сен.', 'январья']

/** TEXT's format codes on a date and time (45688.75 is 2025-01-31 18:00). */
const DATE_CODES = [
  'ДД.ММ.ГГГГ', 'дд.мм.гггг', 'Д.М.ГГ', 'ДДД', 'ДДДД', 'МММ', 'ММММ', 'МММММ', 'ГГГГ', 'ГГ', 'ДД ММММ ГГГГ', 'ММММ ГГГГ', 'Д МММ', 'ч:мм', 'чч:мм:сс',
  'ч:мм:сс AM/PM', 'ч:мм AM/PM', 'ЧЧ:ММ', 'ММ:СС', 'мм:сс', 'ч', 'ДД.ММ.ГГГГ ч:мм', 'ДД.ММ.ГГГГ чч:мм:сс', '[ч]:мм', 'ГГГГ-ММ-ДД', 'дддд', 'ддд', 'мммм',
  'ммм', 'yyyy-mm-dd', 'dd.mm.yyyy', 'h:mm', 'hh:mm:ss', 'd mmmm yyyy', 'YYYY', 'h:mm AM/PM', 'Основной', 'General', 'ДД/ММ/ГГГГ', 'ч:мм:сс,0',
]

/** More date codes, on 2025-01-31 18:07:05: minutes and months by case and place, single letters, AM/PM forms. */
const DATE_CODES_MORE = [
  'м', 'ч:м', 'М', 'ММ', 'ч:ММ', 'ММ:сс', 'М:С', 'Д', 'ДДДДД', 'ГГГ', 'Г', 'Ч', 'с', 'сс,00', 'AM/PM', 'A/P', 'am/pm', 'ч AM/PM', 'ДД.ММ.ГГ ч:мм:сс', 'ДД-МММ-ГГ',
  'ГГГГММДД', 'ч "ч" мм "мин"', 'Ч:мм', 'ЧЧ', 'МММ ГГ', 'm', 'ДД\\.ММ', '"Дата: "ДД.ММ.ГГГГ', 'ч:мм:сс', 'ДД.ММ ч:м:с', 'Д МММ ГГГГ', 'ч:мм:сс,000',
]

/** TEXT's format codes on numbers. */
const NUMBER_CODES = [
  'Основной', 'основной', 'General', '0', '0,00', '# ##0', '# ##0,00', '#,##0.00', '0.00', '0.', '0%', '0,0%', '0.0%', '0,00E+00', '0.00E+00',
  '@', '"руб."0', '0,00" ₽"', '# ##0 ₽', '$#,##0.00', '0,0;-0,0;"ноль"', '0,', '0,0,', '# ##0,0 " тыс."', '0\\.0', '0"."0', '#,##0', '0,000',
  '# ??/??', '#\\ ?/?', '?/?', '# ?/?', '0 ?/8',
]

/** More number codes, on 1234567.891 and −0.5: the point, spaces that group or scale, literals. */
const NUMBER_CODES_MORE = [
  '0.0', '.0', '#.##', '0 . 0', '0 ', '0  ', '# ##0  ', '# ##0 "тыс."', '0,0 ₽', '0 ₽', '0 0', '# ##0,00 ₽', '0,00 ;-0,00 ', '0 %', '0,0 %', '# ##0,0  " млн"',
  '0,00 "р."', '0" "0', '0\\ 0', '#,##0,00', '0.0;0', '0,0.', '0,0.0', 'Основной;Основной', 'Основной "шт."', '"x"0.', '0.%', '0,00-', '(0,00)', '# ##0,', '0,00\\.',
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (group: string, formula: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, formula })

  for (const t of NUMBERS) for (const f of [`VALUE(${q(t)})`, `${q(t)}*1`]) add('number', f)
  for (const t of NUMBERS_MORE) for (const f of [`VALUE(${q(t)})`, `${q(t)}*1`]) add('number', f)
  for (const f of ['SUM("1,5")', 'SUM("1.5")', 'SUM("1 234,5",1)', 'COUNT("1,5","1.5","31.01.2025","abc","1/2/2025","18:30")', 'ABS("-1,5")', '-"1,5"', '"1,5"%', '"1,5"^2', '"ИСТИНА"+0'])
    add('number', f)

  for (const t of DATES) for (const f of [`VALUE(${q(t)})`, `${q(t)}*1`, `DATEVALUE(${q(t)})`, `TIMEVALUE(${q(t)})`, `DAY(${q(t)})`]) add('date', f)
  MONTHS.forEach((forms) => {
    for (const name of new Set(forms)) {
      add('month', `DATEVALUE(${q(`15 ${name} 2025`)})`)
      add('month', `DATEVALUE(${q(`15-${name}-2025`)})`)
    }
    add('month', `MONTH(${q(`1 ${forms[1]}`)})`)
  })
  for (const name of MONTH_EXTRA) add('month', `DATEVALUE(${q(`15 ${name} 2025`)})`)
  for (const t of DATES_MORE) for (const f of [`VALUE(${q(t)})`, `DATEVALUE(${q(t)})`]) add('date', f)
  for (const f of ['EDATE("31.01.2025",1)', 'EOMONTH("15.01.2025",0)', 'WEEKDAY("1.5")', 'YEAR("1.5")', 'MONTH("1.5")', 'YEAR("1.5.25")', 'YEARFRAC("1.1.2025","1.7.2025")', 'NETWORKDAYS("1.1.2025","31.1.2025")', 'WORKDAY("1.1.2025",5,"6.1.2025")', '_xlfn.DAYS("1.3.2025","1.2.2025")', 'DATE(2025,1,"1,5")'])
    add('date', f)

  for (const t of ['ИСТИНА', 'истина', 'Истина', 'ЛОЖЬ', 'ложь', 'TRUE', 'true', 'FALSE', ' ИСТИНА', 'ИСТИНА ', '1', '0', 'да', '']) add('boolean', `IF(${q(t)},1,2)`)
  for (const t of ['ИСТИНА', 'ЛОЖЬ', 'TRUE', 'ложь']) add('boolean', `NOT(${q(t)})`)
  for (const f of ['_xlfn.IFS("ИСТИНА",1,TRUE,2)', '_xlfn.IFS("TRUE",1,TRUE,2)', 'OR("ИСТИНА")', 'AND("ИСТИНА",TRUE)', 'VLOOKUP(1,D!L1:N6,2,"ЛОЖЬ")', 'VLOOKUP(1,D!L1:N6,2,"FALSE")'])
    add('boolean', f)

  for (const x of ['1.5', '-0.5', '0.1', '1234567.5', '1E+21', '1.5E-11', '0.0000000745', '123456789012345678', '(1/3)', 'TRUE', 'FALSE', '(2/3E-15)', '-1.23456789012345E+20'])
    add('text', `"a"&${x}`)
  for (const f of [
    'CONCATENATE(1.5,TRUE)', '_xlfn.CONCAT(-2.5,FALSE)', '_xlfn.TEXTJOIN(1.5,TRUE,"a","b")', '_xlfn.TEXTJOIN(";",TRUE,1.5,TRUE,FALSE)', 'LEN(TRUE)', 'LEN(FALSE)', 'UPPER(1.5)',
    'LOWER(TRUE)', 'EXACT(1.5,"1,5")', '_xlfn.VALUETOTEXT(1.5)', '_xlfn.VALUETOTEXT(TRUE)', '_xlfn.VALUETOTEXT(NA())', '_xlfn.VALUETOTEXT(1/0)', '_xlfn.VALUETOTEXT("a",1)',
    '_xlfn.ARRAYTOTEXT({1.5,2;TRUE,"a"},0)', '_xlfn.ARRAYTOTEXT({1.5,2;TRUE,"a"},1)', '_xlfn.ARRAYTOTEXT({1,#N/A},0)', '_xlfn.ARRAYTOTEXT({1,#N/A},1)', 'HYPERLINK(1.5)', 'TEXT(1.5,"@")',
    'SUBSTITUTE(1.5,",",";")', 'FIND(",",1.5)', 'SEARCH(".",1.5)', 'REPT(-0.25,2)', '_xlfn.ARRAYTOTEXT({"a";"b"},1)', '_xlfn.ARRAYTOTEXT({1.5;-2},0)',
    '_xlfn.ARRAYTOTEXT({#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#NULL!},0)', '_xlfn.VALUETOTEXT(#NUM!)',
  ])
    add('text', f)

  for (const c of DATE_CODES) add('format', `TEXT(45688.75,${q(c)})`)
  for (const c of NUMBER_CODES) for (const v of ['1234.567', '-1234.567', '0.285', '4.34']) add('format', `TEXT(${v},${q(c)})`)
  for (const c of DATE_CODES_MORE) add('format', `TEXT(45688.75+TIME(0,7,5),${q(c)})`)
  for (const c of NUMBER_CODES_MORE) for (const v of ['1234567.891', '-0.5']) add('format', `TEXT(${v},${q(c)})`)
  for (let m = 1; m <= 12; m++) for (const c of ['ММММ', 'МММ', 'ДД ММММ', 'МММММ']) add('format', `TEXT(DATE(2025,${m},15),${q(c)})`)
  for (let d = 0; d < 7; d++) for (const c of ['ДДДД', 'ДДД']) add('format', `TEXT(${45662 + d},${q(c)})`)
  for (const f of ['TEXT("abc","0,00")', 'TEXT("abc","@")', 'TEXT("abc","$#,##0.00")', 'TEXT("abc","0.00")', 'TEXT("1,5","0,00")', 'TEXT("1.5","ДД.ММ")', 'TEXT(TRUE,"0")', 'TEXT(TRUE,"0.00")', 'TEXT(D!A1,"0,00")', 'TEXT(46143.5,"ДД.ММ.ГГГГ ч:мм")'])
    add('format', f)

  for (const [x, d] of [['1234.567', '2'], ['-1234.567', '2'], ['0.5', '0'], ['-0.001', '2'], ['1234567.891', '-3'], ['1E+15', '2'], ['-0.123', '4'], ['0', '0'], ['-1234.5', '0'], ['1234.5', '"1,5"']])
    add('currency', `DOLLAR(${x},${d})`)
  for (const f of ['FIXED(1234567.891,2)', 'FIXED(1234567.891,2,TRUE)', 'FIXED(-0.001,2)', 'FIXED(-1234.5,0)', 'FIXED(1E+15,1)', 'FIXED(1234.5,"1,5")', 'FIXED(1234.5,"1.5")', 'FIXED("1,5",1)'])
    add('currency', f)

  for (const args of [
    '"1 234,5"', '"1.234,5"', '"1.234,5",",","."', '"1,234.5",".",","', '"1,5"', '"1.5"', '"1.5","."', '"1 234.5","."', '"1,5%"', '" 1 234 , 5 "', '"1234,5",","," "',
    '"1.5",".",","', '"1.234,5",","', '"1 234,5",","', '"1,234.5","."', '"1.234.5","."', `"1${NBSP}234,5"`, '"1,5E+3"', '"(1,5)"', '"31.01.2025"', '"18:30"', '"-1,5"',
    '"1.5",","', '"1.5",",","."', '"1 234.5",".",","', '"1.234,5","."',
  ])
    add('numbervalue', `_xlfn.NUMBERVALUE(${args})`)

  for (const f of [
    'IMSUM("1,5+2i")', 'IMSUM("1.5+2i")', 'IMREAL("1,5")', 'IMREAL("1.5")', 'IMREAL("1,5E+3")', 'IMREAL("1 234,5")', 'IMREAL("1,5%")', 'IMREAL("31.01.2025")', 'IMREAL("$1")',
    'COMPLEX(1.5,-2.25)', 'COMPLEX(1E+21,1)', 'IMSUM("1,5e-3i")', 'IMABS("3,5+4i")', 'IMSUB("1,5","2,5i")', 'IMPRODUCT("1,5",2)', 'IMSUM(",5+i")', 'IMREAL("2025-01-31")',
    'IMAGINARY("1+1,5j")', 'IMSUM(1.5,"2,5")', 'IMSUM("1.5+2.5i")', 'IMREAL("1.5+2i")', 'IMAGINARY("1+1.5i")', 'IMSUM("31.01.2025+i")', 'IMSUM("1,5+2.5i")',
    'IMSUM("1 234,5+i")', 'IMSUM("50%+i")', 'IMSUM("1/2+i")', 'IMSUM("1,5E+3+2i")', 'IMSUM("1.5i")', 'IMAGINARY("1.5i")', 'IMREAL("12.5+i")', 'IMSUM("1,5+i")', 'IMREAL("(1,5)")',
    'IMREAL("1 234,5 ₽")', 'IMSUM("1.5-2i")', 'IMREAL("0.5+i")', 'IMREAL("1.13+i")', 'IMSUM("2,5E-1-0,5i")',
  ])
    add('complex', f)

  for (const c of [
    '">1,5"', '">1.5"', '"1,5"', '"=1,5"', '"1.5"', '"<>1,5"', '">=1.5"', '"=1.5"', '">=01.01.2026"', '"<31.12.2025"', '">2025-01-01"', '"=01.05.2026"', '"01.05"', '">1/1/2026"',
    '"ИСТИНА"', '"=ИСТИНА"', '"TRUE"', '"<>ЛОЖЬ"', '"истина"', '"#Н/Д"', '"#N/A"', '"1 234,5"', '"=1 234,5"', '"<2"', '"*,5"', '"1,5*"', '1.5', '46143', 'TRUE', '"<м"', '">a"',
  ])
    add('criteria', `COUNTIF(L!A1:A16,${c})`)
  for (const f of ['SUMIF(L!A1:A16,">1,5")', 'COUNTIFS(L!A1:A16,">=1",L!A1:A16,"<=2,5")', '_xlfn.MAXIFS(L!A1:A16,L!A1:A16,"<01.01.2026")'])
    add('criteria', f)
  for (const [cell, c] of [
    ['A1', '"TRUE"'], ['B1', '"TRUE"'], ['A1', '"=TRUE"'], ['B1', '"=TRUE"'], ['C1', '"ИСТИНА"'], ['B1', '"ИСТИНА"'], ['A1', 'TRUE'], ['C1', 'TRUE'], ['D1', '"<>2"'], ['E1', '"<>2"'],
    ['D1', '"=2"'], ['D1', '2'], ['D1', '">1"'], ['F1', '1234.5'], ['F1', '"<>1234,5"'], ['G1', '46143'], ['G1', '"<>01.05.2026"'], ['H1', '45688'], ['I1', '0.5'], ['J1', '0.015'],
    ['J1', '"1,5%"'], ['I1', '"50%"'], ['H1', '"31.01.2025"'],
  ])
    add('criteria', `COUNTIF(M!${cell},${c})`)
  return out
}
