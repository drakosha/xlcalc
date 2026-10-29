/**
 * The engine's answers to a filter set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/filter/predict.ts <set>` writes `filter/<set>.predicted.json`, each case's AutoFilter after each step
 * and at the end (`filterInEngine`).
 */
import { setModule, writePrediction } from '../stored.js'
import { filterInEngine, type FilterSet } from './cases.js'

const set = process.argv[2]
const { buildFilterCases } = await setModule<FilterSet>('filter', set)
writePrediction(new URL(`./${set}.ts`, import.meta.url).href, set, { cases: Object.fromEntries(buildFilterCases().map((c) => [c.id, filterInEngine(c)])) })
