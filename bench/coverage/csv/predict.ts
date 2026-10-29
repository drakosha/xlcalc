/**
 * The engine's answers to a csv set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/csv/predict.ts <set>` writes `csv/<set>.predicted.json`, for each CSV file the sheet and the
 * cells `readCsv` reads as Excel does (the locale's list separator, formulas taken, the year of today), and for each
 * case the CSV `writeCsv` writes of its book.
 */
import { registerLocale } from '../../../src/index.js'
import { ruRU } from '../../../src/locales/ru-RU.js'
import { csvSet } from '../csv.js'
import { writePrediction } from '../stored.js'
import { engineAnswer } from './cases.js'

registerLocale(ruRU)
const set = process.argv[2]
const { locale, cases, files } = await csvSet(set)
const now = `${new Date().toISOString().slice(0, 10)}T12:00:00Z`
writePrediction(import.meta.url, set, { now, answers: Object.fromEntries(cases.map((c) => [c.id, engineAnswer(c, files.get(c.id)!, locale, now)])) })
console.log(`${set}: ${cases.length} cases predicted`)
