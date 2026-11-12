/**
 * The answers predicted for `cse-review-held.ts` before Excel computed it, by hand (the engine answered otherwise then).
 * A1:A3 3, 1, 2; B1:B3 30, 10, 20.
 */
import type { Val } from '../formulajs.js'

const n = (v: number): Val => ({ t: 'n', v })
const at = (cells: string[], values: Val[]) => Object.fromEntries(values.map((v, i) => [`S!${cells[i]}`, v]))
const G12 = (a: number, b: number) => at(['G1', 'G2'], [n(a), n(b)])
const G56 = (a: number, b: number) => at(['G5', 'G6'], [n(a), n(b)])

export const PREDICTED: Record<string, Record<string, Val>> = {
  'sum-plus-one': G12(11, 21),
  'abs-sum': G12(10, 20),
  'sum-plus-sum': G12(20, 40),
  'max-index-times': G12(20, 40),
  'sum-offset-minus': G12(9, 19),
  'average-plus-sum': G12(70, 80),
  'sum-transpose-plus': G12(11, 21),
  'if-sum-plus': G12(10, 20),
  'sum-plus-one-single': at(['G1'], [n(11)]),
  'sum-inner-plus': G12(32, 32),
  'countif-two-arrays': G56(1, 2),
  'subtotal-function-array': G56(40, 15),
  'countif-criteria-row': at(['G5', 'H5', 'G6', 'H6'], [n(1), n(2), n(1), n(2)]),
  'sumif-criteria-array': G56(30, 30),
  'subtotal-function-longer': at(['G5', 'G6', 'G7'], [n(40), n(15), { t: 'e', v: 'NA' }]),
  'cse-countif-two-arrays': G12(1, 2),
}
