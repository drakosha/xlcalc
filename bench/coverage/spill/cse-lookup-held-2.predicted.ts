/**
 * The answers predicted for `cse-lookup-held-2.ts` before Excel computed it: an argument of Excel's array class
 * (`src/legacy/params.ts`: ROWS, MODE, SUMX2PY2) evaluates the lookup whole over two cells, as SUMPRODUCT's did;
 * INDIRECT selects one reference per position, as OFFSET did, and OFFSET does so through its height too.
 */
import type { Val } from '../formulajs.js'

const n = (...values: number[]) => Object.fromEntries(values.map((v, i) => [`S!G${i + 1}`, { t: 'n', v } as Val]))

export const PREDICTED: Record<string, Record<string, Val>> = {
  'rows-two': n(2, 2),
  'mode-two': n(10, 10),
  'sumx2py2-two': n(500, 500),
  'indirect-two': n(10, 20),
  'indirect-single': n(10),
  'offset-height-two': n(30, 40),
}
