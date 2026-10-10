/**
 * The engine's answers to a set of formatted rows and columns as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/lines/predict.ts <set>` writes `lines/<set>.predicted.json`, for each case what the read cells
 * show after its steps, and for a case without steps the book the reader is to read from the file Excel saves: the
 * book written.
 */
import { writeFileSync } from 'node:fs'
import { entries } from '../xlsx/compare.js'
import { engineCase, type LinesSet } from './cases.js'

const [set] = process.argv.slice(2)
if (!/^[a-z0-9-]+$/.test(set ?? '')) throw new Error(`set name: [a-z0-9-], not ${set}`)
const built = ((await import(`./${set}.js`)) as LinesSet).buildLinesCases()
const out: Record<string, unknown> = {}
for (const c of built.cases) out[c.id] = { looks: engineCase(built, c), ...(c.steps ? {} : { read: Object.fromEntries(entries(built.books[c.book])) }) }
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
console.log(`${set}: ${built.cases.length} cases predicted`)
