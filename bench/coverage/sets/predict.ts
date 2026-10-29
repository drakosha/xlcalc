/**
 * The engine's answers to case sets in en-US as their predictions, written before Excel runs them:
 * `npx tsx bench/coverage/sets/predict.ts <set> …` writes `sets/<set>.predicted.json` (the answer by case id), computed
 * in one book as `test/excel-sets.test.ts` computes the set. After the run the set's `MISSED` lists the ids Excel
 * answered otherwise (`test/excel-predicted.test.ts`).
 */
import { writeFileSync } from 'node:fs'
import { loadSet } from '../../../test/excel-sets.js'
import { localeAnswers } from '../../../test/excel-locale-cases.js'

for (const set of process.argv.slice(2)) {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = await loadSet(set)
  const probes = mod.buildProbes()
  const answers = localeAnswers('en-US', probes, mod.SHEETS)
  writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(Object.fromEntries(probes.map((p, i) => [p.id, answers[i]])), null, 1) + '\n')
  console.log(`${set}: ${probes.length} cases predicted`)
}
