/**
 * The engine's answers to a set of data validation and outlines as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/valid/predict.ts <set>` writes `valid/<set>.predicted.json`, for each case the engine's refusal
 * or each sheet's rules, checks of values and lines over the window after what Excel makes and the operation.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { editStructure } from '../../../src/index.js'
import { bookViews, made, WINDOW, type ValidSet } from './cases.js'

const set = process.argv[2]
const { books, cases, window } = (await setModule<ValidSet>('valid', set)).buildValidCases()
const out: Record<string, unknown> = {}
for (const c of cases) {
  const book = made(books[c.book], c.make)
  const r = c.op ? editStructure(book, c.op) : { ok: true as const, book }
  out[c.id] = r.ok ? bookViews(r.book, window ?? WINDOW) : { refused: r.code }
}
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), out)
console.log(`${set}: ${cases.length} cases predicted`)
