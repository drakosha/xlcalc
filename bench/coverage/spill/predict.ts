/**
 * The engine's answers to a held-out spill set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/spill/predict.ts <set>` writes `spill/<set>.predicted.json` (the non-blank cells of each case's
 * rectangles by `Sheet!A1`, as `test/spill-excel.test.ts` compares them), which the set reads.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { cellsOf } from '../case-book.js'
import type { SpillSet } from './cases.js'
import { computed } from '../../../test/spill-cases.js'
import { shownValue } from '../../../test/spill-shown.js'

const set = process.argv[2]
const out: Record<string, Record<string, unknown>> = {}
for (const c of (await setModule<SpillSet>('spill', set)).buildSpillCases()) {
  const r = computed(c)
  const vals: Record<string, unknown> = {}
  for (const g of c.read) for (const a of cellsOf(g.range).addrs) {
    const v = shownValue(r, g.sheet, a)
    if (v) vals[`${g.sheet}!${a}`] = v
  }
  out[c.id] = vals
}
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), out)
console.log(`${set}: ${Object.keys(out).length} cases predicted`)
