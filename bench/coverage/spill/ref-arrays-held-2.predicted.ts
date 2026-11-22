/**
 * The answers predicted for `ref-arrays-held-2.ts` before Excel computed it, by hand: where the function takes a
 * reference only, one call for each reference and an array of their shape; where it takes an array, the references'
 * top-left values. A legacy array formula takes the references whole under these functions, as under SUBTOTAL.
 * A1:A3 3, 1, 2; B1:B3 30, 10, 20; C1:C3 300, 100, 200; E1 =B1*2, E2 =B2+1, E3 5; J1:K4 k a b a, v 1 2 4; L1:M2 k a, v >1.
 */
import type { Val } from '../formulajs.js'
import { b, e, n, s } from '../vals.js'

const at = (cells: string[], values: Val[]) => Object.fromEntries(values.map((v, i) => [`S!${cells[i]}`, v]))
const G5 = (...values: Val[]) => at(['G5', 'G6', 'G7'], values)
const row5 = (...values: Val[]) => at(['G5', 'H5', 'I5'], values)

export const PREDICTED: Record<string, Record<string, Val>> = {
  'aggregate-9': G5(n(40), n(30)),
  'aggregate-4-row': row5(n(30), n(300)),
  'aggregate-3-indirect': G5(n(2), n(3)),
  'aggregate-two-arrays': G5(n(330), n(110)),
  'aggregate-14-array-form': G5(n(30)),
  'sum-aggregate': G5(n(60)),
  rank: G5(e('NA'), n(1)),
  'rank-avg-ascending': G5(e('NA'), n(2)),
  areas: G5(n(1), n(1)),
  row: G5(n(1), n(2), n(3)),
  column: row5(n(2), n(3)),
  'cell-address': G5(s('$B$1'), s('$B$2')),
  formulatext: G5(s('=B1*2'), s('=B2+1')),
  isformula: G5(b(true), b(true), b(false)),
  'dsum-database': G5(n(1), n(5)),
  'dcount-criteria': row5(n(2), n(1)),
  'rows-values': G5(n(2)),
  'index-values': G5(n(10)),
  isref: G5(b(true), b(true)),
  'cse-aggregate': at(['G1', 'G2'], [n(40), n(30)]),
  'cse-sum-rank': at(['G1'], [n(5)]),
}
