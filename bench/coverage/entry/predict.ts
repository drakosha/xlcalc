/**
 * `parseEntry`'s answers to an entry set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/entry/predict.ts <set>` writes `entry/<set>.predicted.json`, the answer of every case in each of the
 * package's locales, with the day whose year completes a date typed without one.
 */
import { writeFileSync } from 'node:fs'
import { parseEntry } from '../../../src/index.js'
import { TAGS } from '../locales.js'
import type { EntrySet } from './cases.js'

const set = process.argv[2]
const { buildEntryCases } = (await import(`./${set}.js`)) as EntrySet
const now = new Date().toISOString().slice(0, 10)
const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(buildEntryCases().map((c) => [c.id, parseEntry(c.text, { locale: tag, now, ...(c.format !== undefined ? { cell: { format: c.format } } : {}) })]))]))
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify({ now, locales }, null, 1) + '\n')
