/**
 * The engine's answers to a fit set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/fit/predict.ts <set>` writes `fit/<set>.predicted.json`, each case's width and shown texts in each of
 * the package's locales, with the day whose year completes a date typed without one.
 */
import { writeFileSync } from 'node:fs'
import { compute, parseEntry } from '../../../src/index.js'
import { displayRules } from '../../../src/locale/rules.js'
import { fitWidths, typedWidth } from '../../../src/layout/fit.js'
import { parseAddress } from '../../../src/model.js'
import { TAGS } from '../locales.js'
import type { FitSet } from './cases.js'
import { engineFit, type FitEngine } from './engine.js'

const set = process.argv[2]
const { buildFitCases } = (await import(`./${set}.js`)) as FitSet
const now = new Date().toISOString().slice(0, 10)
const clock = { now: `${now}T12:00:00Z`, seed: 1 }

const engine: FitEngine = {
  compute: (book) => compute(book, { clock }),
  parseEntry,
  fit(book, result, col) {
    const c = parseAddress(`${col}1`).col
    const cells = Object.entries(result.sheets[0].cells).map(([a, x]) => [parseAddress(a).row, parseAddress(a).col, x.value] as const)
    return fitWidths(displayRules(book.locale), book.sheets[0], cells, c, c)[col]
  },
  widen: (value, format, tag, column) => typedWidth(displayRules(tag), value, format, column),
}

const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(buildFitCases(tag).map((c) => [c.id, engineFit(engine, c, tag, now)]))]))
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify({ now, locales }, null, 1) + '\n')
