/**
 * The engine's answers to case sets in ru-RU as their predictions, written before Excel ru-RU runs them:
 * `npx tsx bench/coverage/sets/predict-ru.ts <set> …` writes `sets/<set>.ru.predicted.json` (the answer by case id),
 * computed in a book with `locale: 'ru-RU'` as `test/excel-locales-sets.test.ts` computes the set. After the run the
 * set's `RU_MISSED` lists the ids Excel answered otherwise (`test/excel-predicted.test.ts`).
 */
import { predictSets } from './predicting.js'

await predictSets(process.argv.slice(2), 'ru-RU', '.ru')
