/**
 * The engine's answers to case sets as their predictions, written before Excel runs them: `npx tsx
 * bench/coverage/sets/predict.ts <set> …` writes `sets/<set>.predicted.json` (the answer by case id), computed in the
 * book `test/excel-sets.test.ts` computes the set in. After the run the set's `MISSED` lists the ids Excel answered
 * otherwise (`test/excel-predicted.test.ts`).
 */
import { writeFileSync } from 'node:fs'
import { compute } from '../../../src/index.js'
import { CLOCK, toVal } from '../../../test/excel-cases.js'
import { caseBook, loadSet } from '../../../test/excel-sets.js'

for (const set of process.argv.slice(2)) {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = await loadSet(set)
  const probes = mod.buildProbes()
  const cells = compute(caseBook(probes, mod.SHEETS), { clock: CLOCK }).sheets.at(-1)!.cells
  writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(Object.fromEntries(probes.map((p, i) => [p.id, toVal(cells[`A${i + 1}`].value)])), null, 1) + '\n')
  console.log(`${set}: ${probes.length} cases predicted`)
}
