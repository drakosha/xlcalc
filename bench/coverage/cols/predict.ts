/**
 * The hypothesis written before Excel ran the overlapping `<col>` cases: Excel opens each file without repairing it, and a
 * column takes its width and style from the last `<col>` over it in the document (later ones over earlier ones); the
 * file Excel saves has no overlaps. Per case: repaired or not, and for each column read the index of the `<col>` it takes
 * (null where none is over it). `npx tsx bench/coverage/cols/predict.ts`.
 */
import { writePrediction } from '../stored.js'
import { CASES, letters } from './cases.js'

const last = (cols: (typeof CASES)[number]['cols'], n: number) => cols.reduce<number | null>((at, [min, max], i) => (n >= min && n <= max ? i : at), null)
writePrediction(
  import.meta.url,
  'overlap',
  Object.fromEntries(CASES.map((c) => [c.id, { repaired: false, from: Object.fromEntries(letters(c).map((l, i) => [l, last(c.cols, c.read[i])])) }])),
)
