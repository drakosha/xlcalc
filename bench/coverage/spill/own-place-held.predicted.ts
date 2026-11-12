/**
 * The answers predicted for `own-place-held.ts` before Excel computed it: ROW, COLUMN and ROWS read only where their
 * references are, so an array spilling over them is no circular reference; SUBTOTAL, SUMIF and COUNTIF take each
 * reference of an array of them and answer an array, one value each.
 */
import type { Val } from '../formulajs.js'

const at = (cells: string[], values: number[]) => Object.fromEntries(values.map((v, i) => [`S!${cells[i]}`, { t: 'n', v } as Val]))

export const PREDICTED: Record<string, Record<string, Val>> = {
  'row-rows': at(['G1', 'G2'], [1, 2]),
  'row-range': at(['G1', 'G2', 'G3'], [1, 2, 3]),
  'column-range': at(['G1', 'H1', 'I1'], [7, 8, 9]),
  'rows-own': at(['G1', 'G2'], [6, 7]),
  'subtotal-offset': at(['G5', 'G6', 'G7'], [30, 10, 20]),
  'sum-subtotal-offset': at(['G5'], [60]),
  'countif-indirect': at(['G5', 'G6'], [1, 1]),
  'sumif-offset': at(['G5', 'G6'], [30, 20]),
}
