/**
 * Excel en-US cases for date text without a year (`1/31`, `Jan 31`, `31-Jan`), which Excel completes with the
 * current year, and for its neighbours with a year or a time; in VALUE, DATEVALUE, arithmetic, TEXT, date
 * arguments and criteria. Computed in 2026 as dynamic array formulas, in Excel en-US
 * (`test/excel/cases/date-noyear.json`); the tests' clock is in 2026 too.
 */
import type { Probe, SheetCells } from '../probes.js'
import { xl } from './arrays/common.js'

/** G: 2026-01-31 and 2025-01-31 as numbers, the text `1/31`. */
export const SHEETS: SheetCells = { G: { A1: 46053, A2: 45688, A3: '1/31', B1: 'Jan 31' } }

const NUMERIC = ['1/31', '12/31', '1/1', '01/31', '2/28', '2/29', '2/30', '1/32', '12/32', '13/1', '0/1', '1/0', '1/99', '1/100', '1/2025', '12/1999', '1/31/', '1-31', '31-1', '1/31 18:00', '1/31 6 PM', '1/31 25:00']
const NAMED = [
  'Jan 31', 'January 31', 'jan 31', 'JAN 31', '31-Jan', '31 Jan', '31 January', '31-jan', 'Jan-31', 'Jan-25', 'Jan 25', '25-Jan', 'Feb 29', 'Feb 30', 'Sep 31', 'Sept 30', 'Jan 0',
  'Jan 32', 'Jan 99', 'Jan 100', 'Jan 2025', 'January 2025', 'Jan-2025', 'Jan/31', '31/Jan', 'Jan31', '31Jan', 'Jan. 31', 'Jan 31,', 'Jan 31 18:00', '31-Jan 6 PM', 'January 31 6:30 PM',
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (fn: string, formula: string) => out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: xl(formula) })
  for (const t of [...NUMERIC, ...NAMED]) {
    add('VALUE', `VALUE("${t}")`)
    add('DATEVALUE', `DATEVALUE("${t}")`)
  }
  for (const t of ['1/31', 'Jan 31', '31-Jan', '1/31 18:00', '2/29', '1/32']) add('operators', `"${t}"+0`)
  for (const t of ['1/31 18:00', 'Jan 31 6 PM', '1/31']) add('TIMEVALUE', `TIMEVALUE("${t}")`)
  for (const t of ['1/31', 'Jan 31', '31-Jan']) {
    add('YEAR', `YEAR("${t}")`)
    add('TEXT', `TEXT("${t}","yyyy-mm-dd")`)
  }
  add('EDATE', 'EDATE("1/31",1)')
  add('EOMONTH', 'EOMONTH("Jan 31",0)')
  add('DAYS', 'DAYS("2/1","1/31")')
  add('WEEKDAY', 'WEEKDAY("1/31")')
  // criteria: a criterion's text and a cell's text without a year
  for (const c of ['"1/31"', '">=1/31"', '"<Jan 31"', '"=31-Jan"']) add('COUNTIF', `COUNTIF(G!A1:A2,${c})`)
  for (const cell of ['G!A3', 'G!B1']) add('COUNTIF', `COUNTIF(${cell},46053)`)
  add('SUMIF', 'SUMIF(G!A1:A2,"1/31",G!A1:A2)')
  return out
}
