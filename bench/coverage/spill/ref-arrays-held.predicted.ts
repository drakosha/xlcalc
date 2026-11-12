/**
 * The answers predicted for `ref-arrays-held.ts` before Excel computed it, by hand (the engine answered #VALUE! then):
 * the function is called once for each reference of the array and answers an array of its shape; arrays of references
 * in several arguments go together; a legacy array formula takes them whole under these functions, as under N.
 * A1:A3 3, 1, 2; B1:B3 30, 10, 20; C1:C3 300, 100, 200.
 */
import type { Val } from '../formulajs.js'

const n = (v: number): Val => ({ t: 'n', v })
const at = (cells: string[], values: Val[]) => Object.fromEntries(values.map((v, i) => [`S!${cells[i]}`, v]))
const G56 = (a: number, b: number) => at(['G5', 'G6'], [n(a), n(b)])

export const PREDICTED: Record<string, Record<string, Val>> = {
  'averageif-offset': G56(30, 20),
  'averageif-offset-both': G56(300, 200),
  'sumifs-two': G56(300, 200),
  'sumifs-fixed-sum': G56(300, 100),
  'countifs-two': G56(1, 1),
  'maxifs-two': G56(300, 200),
  'countblank-offset': G56(1, 2),
  'subtotal-1': G56(20, 15),
  'subtotal-4-row': at(['G5', 'H5'], [n(30), n(300)]),
  'subtotal-2-indirect': G56(2, 3),
  'subtotal-indirect-sheet': G56(30, 400),
  'subtotal-2d': at(['G5', 'H5', 'G6', 'H6'], [n(30), n(10), n(300), n(100)]),
  'subtotal-ref-error': at(['G5', 'G6'], [n(30), { t: 'e', v: 'REF' }]),
  'subtotal-two-arrays': G56(330, 110),
  'sum-countif-offset': at(['G5'], [n(2)]),
  'sumproduct-subtotal': at(['G5'], [n(3)]),
  'n-offset': G56(30, 10),
  'cse-subtotal-three': at(['G1', 'G2', 'G3'], [n(30), n(10), n(20)]),
  'cse-sumif-two': at(['G1', 'G2'], [n(30), n(20)]),
  'cse-sum-countif-two': at(['G1', 'G2'], [n(2), n(2)]),
  'cse-averageif-single': at(['G1'], [n(50)]),
}
