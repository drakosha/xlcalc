/**
 * Held out after `edit-held`, predicted by `entryText` before Excel ran it: where the formula bar turns a number to
 * scientific notation. `edit-held` showed `100000000000000000000` (21 characters) written whole but `1E-20` in
 * scientific notation and `0.0000001` whole; these numbers, typed in scientific notation, sit on both sides of a length
 * of 21 characters, large and small, with a sign and with 15 significant digits, in both locales' forms.
 */
import { collector, type EntryCase } from './cases.js'

export const edit = true

const NUMBERS: Record<string, string[]> = {
  large: ['1E+21', '9.99999999999999E+20', '1.5E+21', '1E+22', '1.23456789012345E+25', '1E+100', '1E+307', '-1E+20', '-1E+21', '1.23456789012345E+20'],
  small: ['1E-5', '1E-6', '1E-8', '1E-9', '1E-10', '1E-11', '1E-15', '1E-18', '1E-19', '1.5E-18', '1.5E-19', '1.23456789012345E-8', '1.23456789012345E-6', '-1E-19'],
  whole: ['0.000123456789012345', '123456789012345', '1234567890123456', '0.5'],
  // the same numbers in ru-RU's form (a minus before a decimal of the other locale would begin a formula there)
  'large-ru': ['9,99999999999999E+20', '1,5E+21', '1,23456789012345E+25', '1,23456789012345E+20'],
  'small-ru': ['1,5E-18', '1,5E-19', '1,23456789012345E-8', '1,23456789012345E-6'],
  'whole-ru': ['0,000123456789012345', '0,5'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(NUMBERS)) for (const t of texts) add(group, t)
  return out
}
