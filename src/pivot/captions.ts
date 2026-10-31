/**
 * The words a PivotTable shows in its cells in the workbook's language (`Grand Total`, `Row Labels`, `(blank)`, `Sum of
 * Amount`), and those a grouping names its fields and labels its groups with (`Years (Date)`, `Qtr1`, `<1/1/2024`):
 * en-US, and ru-RU as Excel for Windows ru-RU writes them. A PivotTable's own captions replace them.
 */
import type { PivotFunction, PivotGroupBy } from '../contract.js'
import { blankWord, EN_FUNCTIONS, RU_FUNCTIONS, defaultValueName } from '../pivot-words.js'

/** A grouping's words: the fields of dates' periods, the groups' labels (the set period-names). */
export interface GroupWords {
  /** A period's field before its base field's name in brackets (`Years (Date)`). */
  periods: Record<Exclude<PivotGroupBy, 'numbers'>, string>
  months: readonly string[]
  /** A quarter's label before its number (`Qtr1`). */
  quarter: string
  /** A day as the short date (`1/31/2024`), the bounds of groups of dates and of a number of days. */
  date: (y: number, m: number, d: number) => string
  /** A day of the year (`31-Jan`). */
  day: (m: number, d: number) => string
  hour: (h: number) => string
  decimal: string
  /** Between a start of one's own and the blank's word, in the first group of numbers (`<10 or (blank)`). */
  or: string
  blank: string
}

export interface PivotWords {
  grandTotal: string
  rowHeader: string
  columnHeader: string
  values: string
  /** A subtotal's label after its item's (`East Total`). */
  total: string
  blank: string
  all: string
  multiple: string
  /** A value field's name by its function and the field's (`Sum of Amount`). */
  valueName: (fn: PivotFunction, field: string) => string
  /** A subtotal's function after its item's label (`East Count`), and before a value field's name for a grand total (`Total Sum of Amount`). */
  functions: Record<PivotFunction, string>
}

const WORDS: Record<string, PivotWords> = {
  'en-US': {
    grandTotal: 'Grand Total', rowHeader: 'Row Labels', columnHeader: 'Column Labels', values: 'Values', total: 'Total',
    blank: blankWord(), all: '(All)', multiple: '(Multiple Items)', functions: EN_FUNCTIONS,
    valueName: (fn, field) => defaultValueName(fn, field),
  },
  'ru-RU': {
    grandTotal: 'Общий итог', rowHeader: 'Названия строк', columnHeader: 'Названия столбцов', values: 'Значения', total: 'Итог',
    blank: blankWord('ru-RU'), all: '(Все)', multiple: '(несколько элементов)', functions: RU_FUNCTIONS,
    valueName: (fn, field) => defaultValueName(fn, field, 'ru-RU'),
  },
}

const pad = (n: number) => String(n).padStart(2, '0')

const GROUP_WORDS: Record<string, GroupWords> = {
  'en-US': {
    periods: { seconds: 'Seconds', minutes: 'Minutes', hours: 'Hours', days: 'Days', months: 'Months', quarters: 'Quarters', years: 'Years' },
    months: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    quarter: 'Qtr', date: (y, m, d) => `${m}/${d}/${y}`, day: (m, d) => `${d}-${GROUP_WORDS['en-US'].months[m - 1]}`,
    hour: (h) => `${h % 12 === 0 ? 12 : h % 12} ${h < 12 ? 'AM' : 'PM'}`, decimal: '.', or: 'or', blank: blankWord(),
  },
  'ru-RU': {
    periods: { seconds: 'Секунды', minutes: 'Минуты', hours: 'Часы', days: 'Дни', months: 'Месяцы', quarters: 'Кварталы', years: 'Годы' },
    months: ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'],
    quarter: 'Кв-л', date: (y, m, d) => `${pad(d)}.${pad(m)}.${y}`, day: (m, d) => `${pad(d)}.${GROUP_WORDS['ru-RU'].months[m - 1]}`,
    hour: String, decimal: ',', or: 'или', blank: blankWord('ru-RU'),
  },
}

/** A grouping's words in a workbook's locale, en-US's for a locale without its own. */
export const groupWordsOf = (locale: string | undefined): GroupWords => GROUP_WORDS[locale ?? 'en-US'] ?? GROUP_WORDS['en-US']

/** The words of a workbook's locale, en-US's for a locale without its own. */
export const wordsOf = (locale: string | undefined): PivotWords => WORDS[locale ?? 'en-US'] ?? WORDS['en-US']
