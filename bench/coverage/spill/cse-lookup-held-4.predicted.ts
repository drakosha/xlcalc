/**
 * The answers predicted for `cse-lookup-held-4.ts` before Excel computed it, from the idioms' documented purpose: inside
 * a legacy array formula `N(IF(1,{2;3}))` hands INDEX (or XLOOKUP, OFFSET) values it takes whole, and `N(OFFSET(…))`
 * turns OFFSET's references into values, so the enclosing SUM adds every selected value over one cell and over two.
 * Without N around IF, and N over a constant alone, the selection stays positional as for a plain constant.
 */
import type { Val } from '../formulajs.js'

const n = (...values: number[]) => Object.fromEntries(values.map((v, i) => [`S!G${i + 1}`, { t: 'n', v } as Val]))

export const PREDICTED: Record<string, Record<string, Val>> = {
  'index-nif-single': n(30),
  'index-nif-two': n(30, 30),
  'index-tif-single': n(2),
  'index-if-single': n(10),
  'index-n-single': n(10),
  'index-nif-column-single': n(30),
  'offset-n-single': n(30),
  'offset-n-two': n(30, 30),
  'offset-nif-single': n(30),
  'xlookup-nif-single': n(30),
  'index-nif-sequence-single': n(30),
  'vlookup-nif-single': n(30),
}
