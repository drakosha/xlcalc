/** The engine's answers to a select set, written before Excel runs it: `npx tsx bench/coverage/select/predict.ts <set>` writes `select/<set>.predicted.json`. */
import { setModule, writePrediction } from '../stored.js'
import { selectedInBook, type SelectSet } from './cases.js'

const set = process.argv[2]
const { BOOKS, buildSelectCases } = await setModule<SelectSet>('select', set)
const cases = buildSelectCases()
const out = Object.assign({}, ...Object.entries(BOOKS).map(([k, b]) => selectedInBook(b, cases.filter((c) => c.book === k))))
writePrediction(import.meta.url, set, out)
console.log(`${set}: ${cases.length} cases predicted`)
