/**
 * An end set's predictions, written before Excel runs the set: `npx tsx bench/coverage/end/predict.ts <set>` writes
 * `end/<set>.predicted.json`, each case's cell in each of the package's locales, by `getEnd` of the day (`endsInBook`),
 * so a set's committed predictions stay what the engine answered before Excel ran it.
 */
import { fileURLToPath } from 'node:url'
import { TAGS } from '../locales.js'
import { setModule, writeAnswers } from '../stored.js'
import { endsInBook, type EndSet } from './cases.js'

const [set] = process.argv.slice(2)
const { BOOKS, buildEndCases } = await setModule<EndSet>('end', set)
const cases = buildEndCases()
const answers = (tag: string): Record<string, string> => {
  const all = Object.assign({}, ...Object.keys(BOOKS).map((b) => endsInBook(BOOKS[b], cases.filter((c) => c.book === b), tag)))
  return Object.fromEntries(cases.map((c) => [c.id, all[c.id]]))
}
const locales = Object.fromEntries(TAGS.map((tag) => [tag, answers(tag)]))
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), { model: 'getEnd', locales })
