/**
 * Held out after `review-held-2`, where Excel refused `=1E308` and `=1E-320` and took `=1E-309` as 0: whether the bounds
 * are the literal's value or its written exponent. The same values written with other exponents (`0.01E309` is 1E+307,
 * `1000E-312` is 1E-309), exponents between -309 and -320, and a literal below the smallest normal number inside a
 * function; predicted by the engine before Excel ran it.
 */
import { type EntryCase, entryCases } from './cases.js'

const GENERAL: Record<string, string[]> = {
  exp: ['=0.01E309', '=0.1E309', '=100E306', '=1000E-312', '=0.001E-306', '=10E-311', '=1E-310', '=1E-312', '=1E-315', '=1E-318', '=1E-319', '=1E-323', '=1E-308+1', '=ISNUMBER(1E-309)'],
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)
