/**
 * The engine's answers to a fit set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/fit/predict.ts <set>` writes `fit/<set>.predicted.json`, each case's width and shown texts in each of
 * the package's locales, with the day whose year completes a date typed without one.
 */
import { setModule, writePrediction } from '../stored.js'
import { compute, entryWidth, fitColumns, parseEntry } from '../../../src/index.js'
import { TAGS } from '../locales.js'
import type { FitSet } from './cases.js'
import { engineFit, type FitEngine } from './engine.js'

const set = process.argv[2]
const { buildFitCases } = await setModule<FitSet>('fit', set)
const now = new Date().toISOString().slice(0, 10)
const clock = { now: `${now}T12:00:00Z`, seed: 1 }

const engine: FitEngine = {
  compute: (book) => compute(book, { clock }),
  parseEntry,
  fitColumns,
  entryWidth,
}

const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(buildFitCases(tag).map((c) => [c.id, engineFit(engine, c, tag, now)]))]))
writePrediction(import.meta.url, `${set}`, { now, locales })
