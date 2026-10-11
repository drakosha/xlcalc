/**
 * The engine's answers to a hidden set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/hidden/predict.ts <set>` writes `hidden/<set>.predicted.json`, for each case the engine's
 * refusal or each sheet's visibility and the read cells' values, and for a book written as it is the sheet Excel is to
 * show on opening, its first visible one (the writer's active tab).
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { engineAnswer, type HiddenSet } from './cases.js'

const set = process.argv[2]
const built = (await setModule<HiddenSet>('hidden', set)).buildHiddenCases()
const out: Record<string, unknown> = {}
for (const c of built.cases) {
  const opened = c.raw ? undefined : built.books[c.book].sheets.find((s) => !s.hidden)!.name
  out[c.id] = c.raw === 'allHidden' ? { refused: 'every sheet hidden: no book' } : { ...(opened ? { opened } : {}), ...engineAnswer(built, c) }
}
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), out)
console.log(`${set}: ${built.cases.length} cases predicted`)
