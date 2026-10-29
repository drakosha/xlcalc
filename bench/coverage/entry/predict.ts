/**
 * The engine's answers (`engineEntry`) to an entry set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/entry/predict.ts <set>` writes `entry/<set>.predicted.json`, the answer of every case in each of the
 * package's locales, with the day whose year completes a date typed without one. For a set of `edit`, `entryText` of the
 * value each answer holds, by the format it holds, is the predicted text of the formula bar.
 */
import { setModule, writePrediction } from '../stored.js'
import { entryText } from '../../../src/index.js'
import { engineEntry } from './compare.js'
import { TAGS } from '../locales.js'
import type { EntrySet } from './cases.js'

const set = process.argv[2]
const { buildEntryCases, edit, data } = await setModule<EntrySet>('entry', set)
const now = new Date().toISOString().slice(0, 10)
const typed = (tag: string) => buildEntryCases().map((c) => [c.id, engineEntry(c, tag, now, data)] as const)
const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(typed(tag))]))
const shown = (tag: string) =>
  Object.fromEntries(typed(tag).map(([id, r]) => [id, r.ok && r.input ? (r.input.value !== undefined ? entryText(r.input.value as string | number | boolean, { locale: tag, now, ...(r.input.numberFormat !== undefined ? { numberFormat: r.input.numberFormat } : {}) }) : `=${r.input.formula}`) : null]))
writePrediction(import.meta.url, `${set}`, { now, locales, ...(edit ? { edit: Object.fromEntries(TAGS.map((tag) => [tag, shown(tag)])) } : {}) })
