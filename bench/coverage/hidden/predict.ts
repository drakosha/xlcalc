/**
 * The engine's answers to a hidden set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/hidden/predict.ts <set>` writes `hidden/<set>.predicted.json`, for each case the engine's
 * refusal or each sheet's visibility and the read cells' values, and for a book written as it is the sheet Excel is to
 * show on opening, its first visible one (the writer's active tab).
 */
import { writeFileSync } from 'node:fs'
import { engineAnswer, type HiddenSet } from './cases.js'

const set = process.argv[2]
if (!/^[a-z0-9-]+$/.test(set ?? '')) throw new Error(`set name: [a-z0-9-], not ${set}`)
const built = ((await import(`./${set}.js`)) as HiddenSet).buildHiddenCases()
const out: Record<string, unknown> = {}
for (const c of built.cases) {
  const opened = c.raw ? undefined : built.books[c.book].sheets.find((s) => !s.hidden)!.name
  out[c.id] = c.raw === 'allHidden' ? { refused: 'every sheet hidden: no book' } : { ...(opened ? { opened } : {}), ...engineAnswer(built, c) }
}
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
console.log(`${set}: ${built.cases.length} cases predicted`)
