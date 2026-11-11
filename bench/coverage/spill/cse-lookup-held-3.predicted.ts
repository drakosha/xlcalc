/**
 * The answers predicted for `cse-lookup-held-3.ts` before Excel computed it: paired arrays and MODE.SNGL evaluate the
 * lookup whole over two cells as SUMX2PY2 and MODE did; SUMPRODUCT evaluates INDEX whole over two cells and selects
 * over one, as it did XLOOKUP; OFFSET selects under AVERAGE as under SUM; HLOOKUP aggregates as VLOOKUP does, and
 * CHOOSEROWS, a function of dynamic arrays, takes its rows whole.
 */
import type { Val } from '../formulajs.js'

const at = (cells: string[], values: number[]) => Object.fromEntries(values.map((v, i) => [`S!${cells[i]}`, { t: 'n', v } as Val]))
const n = (...values: number[]) => at(['G1', 'G2'], values)

export const PREDICTED: Record<string, Record<string, Val>> = {
  'correl-two': n(1, 1),
  'slope-two': n(10, 10),
  'forecast-two': n(40, 40),
  'mode-sngl-two': n(10, 10),
  'sumproduct-index-two': n(30, 30),
  'sumproduct-index-single': at(['G1'], [10]),
  'average-offset-two': n(10, 20),
  'hlookup-two': at(['G1', 'H1'], [110, 110]),
  'chooserows-two': n(30, 30),
}
