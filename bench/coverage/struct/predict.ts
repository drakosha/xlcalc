/**
 * The engine's answers to a held-out structure set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/struct/predict.ts <set>` writes `struct/<set>.predicted.json`, for each case the engine's
 * refusal, or the look of each sheet after the operation (`LOOK` and the set's `look`, as `test/structure-excel.ts`
 * compares it) and its cells as Excel shows them, in `READ`.
 */
import { setModule, takenBook, writePrediction } from '../stored.js'
import type { StructSet } from './cases.js'
import { LOOK } from '../struct.js'
import { editStructure } from '../../../src/index.js'
import { lookOf } from '../../../test/structure-excel.js'
import { expandSheet } from '../../../test/looks-book.js'

const set = process.argv[2]
const { books, cases, look = [] } = (await setModule<StructSet>('struct', set)).buildStructCases()
const rects = [LOOK, ...look].join(',')
const out: Record<string, unknown> = {}
for (const c of cases) {
  const first = c.before ? editStructure(takenBook(books[c.book]), c.before) : null
  const r = first && !first.ok ? first : editStructure(first ? first.book : takenBook(books[c.book]), c.op)
  out[c.id] = r.ok ? { sheets: Object.fromEntries(r.book.sheets.map((s) => [s.name, { look: lookOf(r.book, expandSheet(s), rects), cells: Object.fromEntries(Object.entries(s.cells).map(([a, x]) => [a, x.formula !== undefined ? `=${x.formula}` : String(x.value)])) }])) } : { refused: r.code }
}
writePrediction(import.meta.url, `${set}`, out)
console.log(`${set}: ${cases.length} cases predicted`)
