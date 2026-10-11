/**
 * The engine's answers to a set of formatted rows and columns as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/lines/predict.ts <set> [run]` writes `lines/<set>.predicted.json` (`<set>.<run>.predicted.json`
 * for a later run), for each case what the read cells show after its steps (taken by a session for a set of `session`,
 * with the book's notes, hyperlinks and rules for a case of `parts`; on the book the reader reads for a bare case), and
 * for a case without steps that writes its book the book the reader is to read from the file Excel saves: the book
 * written.
 */
import { writeFileSync } from 'node:fs'
import { entries } from '../xlsx/compare.js'
import { caseBook, engineCase, partsOf, sessionBook, sessionCase, type LinesSet } from './cases.js'

const [set, run] = process.argv.slice(2)
if (!/^[a-z0-9-]+$/.test(set ?? '') || (run !== undefined && !/^[a-z0-9-]+$/.test(run))) throw new Error(`set and run names: [a-z0-9-], not ${set} ${run}`)
const built = ((await import(`./${set}.js`)) as LinesSet).buildLinesCases()
const out: Record<string, unknown> = {}
const engine = built.session ? sessionCase : engineCase
for (const c of built.cases) {
  const book = await caseBook(built, c)
  out[c.id] = { looks: engine(c, book), ...(c.parts ? { parts: partsOf(sessionBook(c, book)) } : {}), ...(c.steps || c.bare ? {} : { read: Object.fromEntries(entries(built.books[c.book])) }) }
}
writeFileSync(new URL(`./${set}${run ? `.${run}` : ''}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
console.log(`${set}: ${built.cases.length} cases predicted`)
