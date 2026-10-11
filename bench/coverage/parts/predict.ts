/**
 * The engine's answers to a parts set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/parts/predict.ts <set>` writes `parts/<set>.predicted.json`, for each case the engine's refusal
 * or each sheet's notes, hyperlinks and frozen panes after the operation (what Excel makes is given as made, with its
 * author unknown: `?`).
 */
import { writeFileSync } from 'node:fs'
import { editStructure } from '../../../src/index.js'
import { bookParts, made, type PartsSet } from './cases.js'

const set = process.argv[2]
if (!/^[a-z0-9-]+$/.test(set ?? '')) throw new Error(`set name: [a-z0-9-], not ${set}`)
const { books, cases, threads } = ((await import(`./${set}.js`)) as PartsSet).buildPartsCases()
const out: Record<string, unknown> = {}
for (const c of cases) {
  const book = made(books[c.book], c.make, '?', threads?.[c.book])
  const r = c.op ? editStructure(book, c.op) : { ok: true as const, book }
  out[c.id] = r.ok ? Object.fromEntries(r.book.sheets.map((s) => [s.name, bookParts(s)])) : { refused: r.code }
}
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
console.log(`${set}: ${cases.length} cases predicted`)
