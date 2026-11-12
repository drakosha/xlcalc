/**
 * The answers predicted for `cse-lookup-held-5.ts` before Excel computed it, from the idioms' documented purpose: SUBTOTAL,
 * SUMIF, COUNTIF and N take every reference OFFSET or INDIRECT returns for an array of shifts or texts, one count or sum
 * each, and the enclosing aggregation adds them; INDEX over ROW(…) and over SMALL(IF(…),ROW(…)) selects by position.
 */
import type { Val } from '../formulajs.js'

const n = (...values: number[]) => Object.fromEntries(values.map((v, i) => [`S!G${i + 1}`, { t: 'n', v } as Val]))

export const PREDICTED: Record<string, Record<string, Val>> = {
  'subtotal-offset-single': n(60),
  'subtotal-offset-two': n(60, 60),
  'subtotal-offset-rows-single': n(3),
  'max-subtotal-offset-single': n(30),
  'sumif-offset-single': n(50),
  'countif-indirect-single': n(2),
  'n-indirect-single': n(30),
  'index-row-single': n(10),
  'index-small-if-two': n(30, 20),
}
