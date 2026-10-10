/**
 * Held out, after `sheet-quotes-held`: the names its R1C1 cases left open. There Excel wrote `R0` and `R1048577`
 * unquoted but `R2D2`, `C3PO` and `RC1C` quoted, so it reads an R1C1 reference at the start of a name; these names tell
 * whether a reference without a number at the start counts (`Region`, `Cost`), how far the reference reads (`R1C0`,
 * `C16385x`, `RCC1`) and that A1 cells and booleans count only as the whole name (`AB12x`, `TRUEx`). `PREDICTED` was
 * committed before Excel ran the set.
 */
import { quoteCases } from './sheet-quotes-held.js'
import type { TableCase } from './cases.js'

export const local = true

export const NAMES = [
  'Region', 'cost', 'Rate', 'Comments', 'Rx', 'Cx', 'RCx', 'R1x', 'C1x', 'R1_x', 'C1.x', 'R1C1x', 'R1048576x', 'C16384x', 'C16385x', 'R0x', 'R00', 'R01',
  'C0', 'CR', 'CC', 'RR', 'RCR', 'RCC', 'R1R', 'C1C1', 'R1C0', 'R1C16385', 'R1C16384', 'R1048576C16384', 'R1048577C1', 'RC0', 'R1C', 'RCC1', 'Rc1', 'R_',
  'C_', 'R.', 'C.', 'RC_', 'R1.', 'Sheet_R1', 'xR1', 'Ř1', 'Р1', 'С1', 'К', 'AB12x', 'A1_', 'A1.x', 'TRUEx', 'FALSE1', 'TRUE_', '_2024',
]

/** The names predicted in quotes; the others were predicted without. */
const QUOTED = ['R1x', 'C1x', 'R1_x', 'C1.x', 'R1C1x', 'R1048576x', 'C16384x', 'R01', 'R1R', 'C1C1', 'R1C16384', 'R1048576C16384', 'R1C', 'Rc1', 'R1.']

/**
 * The rule predicted from `sheet-quotes-held`: quoted when the name begins with an R1C1 reference read as far as it goes
 * (`R` and its number, then `C` and its number, numbers of the grid) that holds a number or is the whole name; a name
 * beginning with `R` or `C` and a letter otherwise (`Region`) stays unquoted.
 */
export const PREDICTED: Readonly<Record<string, string>> = Object.fromEntries(NAMES.map((name) => [name, `=${QUOTED.includes(name) ? `'${name}'` : name}!A1`]))

export const buildTableCases = (): TableCase[] => quoteCases(NAMES)
