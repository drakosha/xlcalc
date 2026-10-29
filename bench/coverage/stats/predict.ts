/**
 * The engine's answers to a stats set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/stats/predict.ts <set>` writes `stats/<set>.predicted.json`, for each case the rows of
 * `Session.getWorkbookStatistics` as the window labels them.
 */
import { setModule, writePrediction } from '../stored.js'
import { engineStats, type StatsSet } from './cases.js'

const set = process.argv[2]
const cases = (await setModule<StatsSet>('stats', set)).buildStatsCases()
const out: Record<string, unknown> = {}
for (const c of cases) out[c.id] = await engineStats(c)
writePrediction(import.meta.url, set, out)
console.log(`${set}: ${cases.length} cases predicted`)
