/**
 * Held out after `review-held`, where Excel refused every formula with a literal past the largest number or below the
 * smallest (`=1E400`, `=1.8E308`, `=1E308*10`, `=1E-400`): where the bounds lie. Literals just below and at 1E+308,
 * the largest number Excel keeps, the smallest normal number and the subnormal ones below it, inside formulas and as an
 * argument; predicted by the engine before Excel ran it.
 */
import { type EntryCase, entryCases } from './cases.js'

const GENERAL: Record<string, string[]> = {
  big: ['=1E307', '=9.99999999999999E307', '=9.999999999999999E307', '=1E308', '=1.7976931348623157E308', '=SUM(1E308,1)', '=1E307*10'],
  small: ['=1E-307', '=2.2250738585072014E-308', '=2.2E-308', '=1E-308', '=1E-309', '=1E-320', '=4.9E-324', '=1E-324', '=0E-400', '=0.0E400'],
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)
