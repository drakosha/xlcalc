/**
 * Excel en-US cases for time text read by VALUE (and arithmetic) and by date and time arguments (TIMEVALUE,
 * HOUR, DATEVALUE): AM and PM with and without a space, hours alone, hours past 24 and past 9999, minutes and
 * seconds past 59, spaces, a date before the time. Whether the two readings are one rule decides whether
 * `dates.ts` keeps two time readers. Computed as dynamic array formulas, in Excel en-US
 * (`test/excel/cases/time-text.json`).
 */
import type { Probe } from '../probes.js'

const TIMES = [
  '6 PM', '6PM', '6pm', '6 pm', '06 PM', '6  PM', '6:30PM', '6:30pm', '6:30 PM', '6:30  PM', '6:30:15PM', '6:30:15 PM', '6:30 AM', '6:30 am', '6:30 Pm', '6:30AM',
  '18:30', '18:30:15.5', '0:00:00.5', '25:00', '24:00', '10:60', '6:60', '6:30:60', '1:99', '1:2:3', '0:0', '0 PM', '12 AM', '12 PM', '13 PM', '13:00 PM', '12:60 PM',
  '6 P', '6:30 P', '6 :30', '6: 30', '6:30 ', '  6:30', '123:00', '1234:00', '9999:00', '9999:59:59', '10000:00', '12345:00', '00123:00', '01234:00', '6.5 PM',
  '6:30 A.M.', '18:30 PM', ':30', '6:', '6:30:', '6::30', '6:30:15.5 PM', '18:30:15,5',
]
const DATE_TIMES = [
  '1/31/2025 6PM', '1/31/2025 6 PM', '1/31/2025 6:30pm', '1/31/2025 25:00', '1/31/2025 9999:00', '1/31/2025 10000:00', '1/31/2025  18:00', '1/31/2025 18:00:00.5',
  '2025-01-31T18:00', '2025-01-31 18:00', '2025-01-31 6PM', '31-Jan-2025 6 PM', 'Jan 31, 2025 18:00', '18:00 1/31/2025',
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (fn: string, formula: string) => out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula })
  for (const t of TIMES) {
    add('VALUE', `VALUE("${t}")`)
    add('TIMEVALUE', `TIMEVALUE("${t}")`)
    add('HOUR', `HOUR("${t}")`)
  }
  for (const t of DATE_TIMES) {
    add('VALUE', `VALUE("${t}")`)
    add('DATEVALUE', `DATEVALUE("${t}")`)
    add('TIMEVALUE', `TIMEVALUE("${t}")`)
  }
  return out
}
