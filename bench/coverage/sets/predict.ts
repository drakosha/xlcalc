/**
 * The engine's answers to case sets in en-US as their predictions, written before Excel runs them:
 * `npx tsx bench/coverage/sets/predict.ts <set> …` writes `sets/<set>.predicted.json`, computed in the set's book as
 * `test/excel-sets.test.ts` computes it; the set reads it as its `PREDICTED`, and after the run its `MISSED` lists the ids
 * Excel answered otherwise (`test/excel-predicted.test.ts`).
 */
import { predictSets } from './predicting.js'

await predictSets(process.argv.slice(2), 'en-US', '')
