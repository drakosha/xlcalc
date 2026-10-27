/**
 * The engine's answers to a set of protection as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/protect/predict.ts <set>` writes `protect/<set>.predicted.json`, for each case what the object
 * model is to read after its steps (`engineCase`: every sheet's protection, the read cells' locked and hidden).
 */
import { setModule, writePrediction } from '../stored.js'
import { engineCase, type ProtectSet } from './cases.js'

const set = process.argv[2]
const built = (await setModule<ProtectSet>('protect', set)).buildProtectCases()
const out: Record<string, unknown> = {}
for (const c of built.cases) out[c.id] = await engineCase(built, c)
writePrediction(import.meta.url, set, out)
console.log(`${set}: ${built.cases.length} cases predicted`)
