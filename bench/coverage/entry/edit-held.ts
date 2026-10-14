/**
 * Held out, predicted by `entryText` before Excel ran it: the text the formula bar opens a cell with (`Formula`,
 * `FormulaLocal`, after the prefix character), for values typed in each form `parseEntry` reads: numbers short and long,
 * small and past 15 digits, percents, money, fractions, dates with and without a time, times, booleans, text after an
 * apostrophe, and numbers typed into cells formatted before as dates, times, money, fractions, scientific and text.
 * Both locales' forms are in the set, which runs in each.
 */
import { type EntryCase, entryCases } from './cases.js'

export const edit = true

const GENERAL: Record<string, string[]> = {
  num: ['1234.5', '1234567.891', '(0.5)', '-5', '0.0000001', '1E-20', '123456789012345678', '12345678901234.5', '100000000000000000000', '1E+20', '1.5E+3', '0.1', '3.14159265358979', '0'],
  'num-ru': ['1234,5', '1234567,891', '(0,5)', '0,0000001', '1,5E+3', '3,14159265358979'],
  pct: ['12%', '12.5%', '-5%', '0.5%', '1234%', '12,5%', '0,5%'],
  money: ['$1,234.50', '$5', '($5)', '1 234,50 ₽', '5 ₽'],
  frac: ['1 1/2', '0 1/3', '2 3/4'],
  date: ['1/31/2025', '1/31/2025 18:00', '1/31/2025 6:30 PM', '31-Jan-2025', 'Jan 2025', '2025-01-31', '12/31/9999'],
  'date-ru': ['31.01.2025', '31.01.2025 18:00', '31 янв 2025', '31.12.9999'],
  time: ['18:30', '6:30 PM', '18:30:15', '25:30', '0:00', '12:00 AM'],
  bool: ['TRUE', 'false', 'ИСТИНА', 'ложь'],
  text: ["'007", "'1.5", "'1,5", "'=1+2", "'TRUE", "'ИСТИНА", "'abc", 'abc', "'12%", "'1/31/2025", "'31.01.2025"],
}

/**
 * Numbers typed into a cell with a format set before: the format the cell keeps. None goes into a percent cell, where
 * `Formula2Local` takes a number otherwise than typing does (`test/entry-known.ts`); text with an apostrophe does.
 */
const FORMATTED: [string, string | number, string[]][] = [
  ['date-cell', 14, ['45688', '45688.75', '45688,75', '0', '-1']],
  ['long-date-cell', 'd mmmm yyyy', ['45688']],
  ['time-cell', 'h:mm', ['0.75', '0,75', '1.75']],
  ['elapsed-cell', '[h]:mm:ss', ['1.5']],
  ['money-cell', '$#,##0.00', ['1234.5']],
  ['frac-cell', '# ?/?', ['0.5']],
  ['sci-cell', '0.00E+00', ['1234.5']],
  ['text-cell', '@', ['007', '1.5']],
  ['grouped-cell', '#,##0.00', ['1234567.891']],
  ['month-cell', 'mmm-yy', ['45688.5']],
  ['pct-text', '0%', ["'5"]],
]

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL, FORMATTED)
