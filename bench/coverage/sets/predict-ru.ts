/**
 * The engine's answers to case sets in ru-RU as their predictions, written before Excel ru-RU runs them:
 * `npx tsx bench/coverage/sets/predict-ru.ts <set> …` writes `sets/<set>.ru.predicted.json` (the answer by case id),
 * computed in a book with `locale: 'ru-RU'` as `test/excel-locales-sets.test.ts` computes the set. After the run the
 * set's `RU_MISSED` lists the ids Excel answered otherwise (`test/excel-predicted.test.ts`).
 */
import { writeFileSync } from 'node:fs'
import { loadSet } from '../../../test/excel-sets.js'
import { localeAnswers } from '../../../test/excel-locale-cases.js'

for (const set of process.argv.slice(2)) {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = await loadSet(set)
  const probes = mod.buildProbes()
  const answers = localeAnswers('ru-RU', probes, mod.SHEETS)
  writeFileSync(new URL(`./${set}.ru.predicted.json`, import.meta.url), JSON.stringify(Object.fromEntries(probes.map((p, i) => [p.id, answers[i]])), null, 1) + '\n')
  console.log(`${set}: ${probes.length} cases predicted in ru-RU`)
}
