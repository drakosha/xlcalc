/**
 * Codes with digit placeholders and a bare letter of a date (`#,##0 m`), as a generator writes a unit: what Excel shows of
 * them, for `isDateFormat`. Predicted before Excel showed them: a digit placeholder with a date's letter is no format of
 * Excel's, so the file is refused or the cell shows General; if it shows, `m` is no month. `0 samples` and a quoted
 * letter are controls, whose letters read as text or are text.
 */
import type { DisplayCase } from './cases.js'

const CODES = ['#,##0 m', '0 m', '0.0 m', '0 d', '0 y', '0 h', '0 s', '0.0 h', '0 samples', '0 "m"', '0 \\m', 'mm 0', 'd 0', '0 yy', '[Red]0 m', '0 m;0', 'dd.mm.yyyy', '[h]:mm']

export function buildDisplayCases(): DisplayCase[] {
  return CODES.flatMap((format, i) =>
    [1234.5, 45688.75].map((value, k) => ({ id: `c${i + 1}#${k + 1}`, cell: { value, format }, width: 30 })),
  )
}
