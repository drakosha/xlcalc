/** Case sets' predictions: the engine's answers in a locale, written as `sets/<set><suffix>.predicted.json` (the answer by case id). */
import { writeFileSync } from 'node:fs'
import { loadSet } from '../../../test/excel-sets.js'
import { localeAnswers } from '../../../test/excel-locale-cases.js'

export async function predictSets(sets: readonly string[], tag: string, suffix: string): Promise<void> {
  for (const set of sets) {
    if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
    const mod = await loadSet(set)
    const probes = mod.buildProbes()
    const answers = localeAnswers(tag, probes, mod.SHEETS)
    writeFileSync(new URL(`./${set}${suffix}.predicted.json`, import.meta.url), JSON.stringify(Object.fromEntries(probes.map((p, i) => [p.id, answers[i]])), null, 1) + '\n')
    console.log(`${set}: ${probes.length} cases predicted in ${tag}`)
  }
}
