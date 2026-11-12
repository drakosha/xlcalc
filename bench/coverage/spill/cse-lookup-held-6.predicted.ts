/**
 * The answers predicted for `cse-lookup-held-6.ts` before Excel computed it: N or T over IF or over OFFSET/INDIRECT hand
 * every value or reference on whole; SUBTOTAL, COUNTIF and SUMIF take each reference of an array of them; INDIRECT over
 * an array of texts without N selects by position like INDEX; extraction and multi-criteria lookups select as Excel
 * documents them; ROW and ROWS over the formula's own cell read only its place, so there is no circular reference.
 */
import type { Val } from '../formulajs.js'

const n = (...values: number[]) => Object.fromEntries(values.map((v, i) => [`S!G${i + 1}`, { t: 'n', v } as Val]))

export const PREDICTED: Record<string, Record<string, Val>> = {
  'nif-true-single': n(30),
  'nif-row-single': n(30),
  'max-nif-single': n(30),
  'average-n-offset-single': n(20),
  'n-offset-two': n(60, 60),
  'subtotal-offset-rows-two': n(60, 60),
  'countif-offset-single': n(2),
  'sumif-indirect-single': n(50),
  'indirect-concat-single': n(30),
  'n-indirect-concat-single': n(40),
  'large-row-indirect-single': n(50),
  'small-if-constant-two': n(30, 20),
  'small-if-iferror-three': n(30, 20, -1),
  'small-if-rows-single': n(30),
  'index-match-keys-two': n(10, 20),
  'index-match-criteria-single': n(20),
  'sum-index-columns-single': n(10),
  'sum-index-row-zero-single': n(330),
  'row-own-single': n(1),
  'rows-own-single': n(2),
  'row-own-two': n(1, 2),
  'sum-row-own-single': n(6),
}
