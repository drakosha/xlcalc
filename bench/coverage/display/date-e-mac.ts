/** The set `date-e-held` without the codes Excel for Windows refuses a file for (`E`, `EE`): the Mac's step stops at a refused book. */
import type { DisplayCase } from './cases.js'
import { buildDisplayCases as held } from './date-e-held.js'

export function buildDisplayCases(): DisplayCase[] {
  return held().filter((c) => c.cell.format !== 'E' && c.cell.format !== 'EE')
}
