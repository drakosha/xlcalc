/**
 * Excel cases for date, time and number text with spaces before or after it, in each date function, in numeric
 * arguments and operators, and in VALUE, DATEVALUE and TIMEVALUE: which of them take leading spaces before a date
 * or a time (the `time-text` set found that HOUR and DAY do not, VALUE does). The text is read alike in both
 * locales (`2025-01-31`, `6:30`, `5`). Computed as dynamic array formulas, in Excel en-US
 * (`test/excel/cases/lead-spaces.json`) and ru-RU (`test/excel/cases/ru-RU/lead-spaces.json`).
 */
import type { Probe, SheetCells } from '../probes.js'
import { xl } from './arrays/common.js'

/** G: cell text with leading spaces, for criteria that read a cell's text as a number. */
export const SHEETS: SheetCells = { G: { A1: ' 2025-01-31', A2: ' 5', A3: ' 6:30', A4: '2025-01-31 ' } }

const D = ['" 2025-01-31"', '"2025-01-31 "', '"  2025-01-31"']
const T = ['" 6:30"', '"6:30 "']

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (fn: string, formula: string) => out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: xl(formula) })
  for (const fn of ['DAY', 'MONTH', 'YEAR', 'WEEKDAY', 'ISOWEEKNUM', 'WEEKNUM', 'DATEVALUE', 'VALUE', 'ABS', 'N', 'INT']) for (const d of D) add(fn, `${fn}(${d})`)
  for (const fn of ['HOUR', 'MINUTE', 'SECOND', 'TIMEVALUE', 'VALUE', 'ABS']) for (const t of T) add(fn, `${fn}(${t})`)
  for (const d of D.slice(0, 2)) {
    add('DAYS', `DAYS(${d},"2025-01-01")`)
    add('DAYS', `DAYS("2025-12-31",${d})`)
    add('DATEDIF', `DATEDIF(${d},"2025-12-31","d")`)
    add('DAYS360', `DAYS360(${d},"2025-12-31")`)
    add('EDATE', `EDATE(${d},1)`)
    add('EOMONTH', `EOMONTH(${d},0)`)
    add('NETWORKDAYS', `NETWORKDAYS(${d},"2025-12-31")`)
    add('WORKDAY', `WORKDAY(${d},1)`)
    add('YEARFRAC', `YEARFRAC(${d},"2025-12-31")`)
    add('operators', `${d}+0`)
    add('operators', `${d}=45688`)
    add('TEXT', `TEXT(${d},"0")`)
    add('DATE', `DATE(2025,1,${d.replace('2025-01-31', '31')})`)
  }
  for (const t of [...T, '" 5"', '"5 "', '"  5"', '" 1E+3"', '" (5)"', '" 50%"']) add('operators', `${t}+0`)
  for (const t of ['" 5"', '" (5)"', '" 50%"']) add('ABS', `ABS(${t})`)
  add('TIMEVALUE', 'TIMEVALUE(" 2025-01-31 6:30")')
  add('COUNTIF', 'COUNTIF(D!A6," 2025-01-31")')
  add('COUNTIF', 'COUNTIF(D!A6,">= 2025-01-31")')
  for (const [cell, crit] of [['G!A1', '45688'], ['G!A2', '5'], ['G!A3', 'TIME(6,30,0)'], ['G!A4', '45688'], ['G!A1', '">=45688"'], ['G!A2', '">4"']]) add('COUNTIF', `COUNTIF(${cell},${crit})`)
  return out
}
