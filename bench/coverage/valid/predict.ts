/**
 * The engine's answers to a set of data validation and outlines as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/valid/predict.ts <set>` writes `valid/<set>.predicted.json`, for each case the engine's refusal
 * or each sheet's rules, checks of values and lines over the window after what Excel makes and the operation; a raw
 * case on the book the reader reads from its file, with what the reader reports.
 */
import { setModule, writePrediction } from '../stored.js'
import { editStructure } from '../../../src/index.js'
import { readXlsx } from '../../../src/xlsx.js'
import { bookViews, made, validCaseFile, WINDOW, type ValidSet } from './cases.js'

const set = process.argv[2]
const built = (await setModule<ValidSet>('valid', set)).buildValidCases()
const { books, cases, window } = built
const out: Record<string, unknown> = {}
for (const c of cases) {
  const read = c.raw ? await readXlsx(await validCaseFile(built, c)) : undefined
  const book = made(read?.book ?? books[c.book], c.make)
  const r = c.op ? editStructure(book, c.op) : { ok: true as const, book }
  const issues = read?.issues.map((i) => `${i.code} ${i.sheet ?? ''} ${(i.cells ?? []).join(' ')}: ${i.message}`)
  out[c.id] = r.ok ? (issues ? { views: bookViews(r.book, window ?? WINDOW), issues } : bookViews(r.book, window ?? WINDOW)) : { refused: r.code }
}
writePrediction(import.meta.url, `${set}`, out)
console.log(`${set}: ${cases.length} cases predicted`)
