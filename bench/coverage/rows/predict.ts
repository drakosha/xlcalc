/**
 * The engine's answers to a rows set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/rows/predict.ts <set>` writes `rows/<set>.predicted.json`, each case's heights and wrapping in each of
 * the package's locales, with the day whose year completes a date typed without one.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { compute, parseEntry } from '../../../src/index.js'
import { fitRows } from '../../../src/layout/heights.js'
import { TAGS } from '../locales.js'
import type { RowSet } from './cases.js'
import { engineRows, type RowsEngine } from './engine.js'

const set = process.argv[2]
const { buildRowCases } = await setModule<RowSet>('rows', set)
const now = new Date().toISOString().slice(0, 10)
const clock = { now: `${now}T12:00:00Z`, seed: 1 }

const engine: RowsEngine = { compute: (book) => compute(book, { clock }), parseEntry, fitRows }

const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(buildRowCases(tag).map((c) => [c.id, engineRows(engine, c, tag, now)]))]))
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), { now, locales })
