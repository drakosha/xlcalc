/**
 * Held out: SUBTOTAL over an array of references (OFFSET or INDIRECT over an array) with rows hidden by hand, the idiom
 * that counts or adds the visible rows one by one (`SUMPRODUCT(SUBTOTAL(103,OFFSET(…)),…)`). Predicted by hand before
 * Excel ran (`ref-arrays-hidden-held.predicted.json`): SUBTOTAL is called once for each reference, 101–111 leave out
 * the hidden rows, 1–11 keep them. S!A1:A6 1, 2, 4, 8, 16, 32; rows 3 and 5 hidden.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { WorkbookCase } from './cases.js'
import { formulaCell as f } from '../stored.js'

const ALL = 'OFFSET(S!A1,{0;1;2;3;4;5},0)'
const PAIRS = 'OFFSET(S!A1,{0;2;4},0,2)'

const FS = [
  `SUM(SUBTOTAL(109,${ALL}))`,
  `SUM(SUBTOTAL(9,${ALL}))`,
  `SUM(SUBTOTAL(103,${ALL}))`,
  `SUM(SUBTOTAL(3,${ALL}))`,
  `SUM(SUBTOTAL(101,${PAIRS}))`,
  `MAX(SUBTOTAL(104,${PAIRS}))`,
  `SUM(SUBTOTAL(102,${PAIRS}))`,
  `SUMPRODUCT(SUBTOTAL(103,${ALL}),S!A1:A6)`,
  `SUM(SUBTOTAL(105,${PAIRS}))`,
  `SUM(SUBTOTAL(106,${PAIRS}))`,
  `SUM(SUBTOTAL(107,${PAIRS}))`,
  'SUM(SUBTOTAL(109,OFFSET(S!A1,ROW(S!A1:A6)-1,0)))',
  `SUM(SUBTOTAL(110,${PAIRS}))`,
  `SUM(SUBTOTAL(108,${PAIRS}))`,
  'SUM(SUBTOTAL(9,INDIRECT({"S!A1:A3";"S!A4:A6"})))',
  'SUM(SUBTOTAL(109,INDIRECT({"S!A1:A3";"S!A4:A6"})))',
]

const data = { name: 'S', cells: Object.fromEntries([1, 2, 4, 8, 16, 32].map((v, i) => [`A${i + 1}`, { value: v }])) }
const book = { sheets: [data, { name: 'F', cells: Object.fromEntries(FS.map((x, i) => [`A${i + 1}`, f(x)])) }] }
const read = [{ sheet: 'F', range: `A1:A${FS.length}` }]

const PREDICTED = JSON.parse(readFileSync(new URL('./ref-arrays-hidden-held.predicted.json', import.meta.url), 'utf8')) as Record<string, Record<string, Val>[]>

export function buildWorkbookCases(): WorkbookCase[] {
  const cases: WorkbookCase[] = [{ id: 'manual', about: 'rows 3 and 5 hidden by hand', book, read, file: { sheets: { S: { hiddenRows: ['3', '5'] } } } }]
  return cases.map((c) => (PREDICTED[c.id] ? { ...c, predicted: PREDICTED[c.id] } : c))
}
