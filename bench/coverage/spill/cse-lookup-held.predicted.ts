/**
 * The answers predicted for `cse-lookup-held.ts` before Excel computed it. Other aggregations are expected to follow
 * SUM's measured rules: the selector picks one lookup per output position unless an arithmetic operator in a multi-cell
 * area evaluates the whole array first. OFFSET is expected to select like INDEX, CHOOSE and LOOKUP to aggregate like
 * VLOOKUP, and a union of cells as a key to be refused. The engine's answers for AVERAGE, MAX, COUNT, SUMPRODUCT,
 * OFFSET, CHOOSE and LOOKUP over G1:G2 were seen before these predictions were written.
 */
import type { Val } from '../formulajs.js'

const n = (...values: number[]) => Object.fromEntries(values.map((v, i) => [`S!G${i + 1}`, { t: 'n', v } as Val]))
const e = (...codes: string[]) => Object.fromEntries(codes.map((v, i) => [`S!G${i + 1}`, { t: 'e', v } as Val]))

export const PREDICTED: Record<string, Record<string, Val>> = {
  'average-two': n(10, 20),
  'average-single': n(10),
  'max-two': n(10, 20),
  'min-two': n(10, 20),
  'count-two': n(1, 1),
  'count-single': n(1),
  'product-two': n(10, 20),
  'sumproduct-two': n(10, 20),
  'average-index-two': n(10, 20),
  'index-xmatch-two': n(10, 20),
  'max-plus-one-two': n(21, 21),
  'offset-two': n(10, 20),
  'offset-single': n(10),
  'choose-two': n(30, 30),
  'lookup-two': n(30, 30),
  'union-key-two': e('VALUE', 'VALUE'),
}
