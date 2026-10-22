/**
 * Held out, predicted by the engine before Excel ran it: formulas with a number literal past the largest number
 * (`=1E400`, negated, inside a function, past it by a little) and the smallest (`=1E-400`); the leap day typed without
 * a year in a year that has none (`2/29` read as February 2029), beside a day past the month's end (`2/30`), a month
 * and a year (`5/2026`) and other spellings of the leap day.
 */
import { type EntryCase, entryCases } from './cases.js'

const GENERAL: Record<string, string[]> = {
  lit: ['=1E400', '=-1E400', '=1E+309', '=ISNUMBER(1E400)', '=1.8E308', '=1E-400', '=1E308*10'],
  leap: ['2/29', '2/30', '5/2026', 'Feb 29', '2-29', '2/28', '4/31'],
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)
