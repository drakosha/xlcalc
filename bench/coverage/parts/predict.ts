/**
 * The engine's answers to a parts set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/parts/predict.ts <set>` writes `parts/<set>.predicted.json`, for each case the engine's refusal
 * or each sheet's notes, hyperlinks and frozen panes after the operation (what Excel makes is given as made, with its
 * author unknown: `?`); a raw case on the book the reader reads from its file, with what the reader reports.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { editStructure } from '../../../src/index.js'
import { readXlsx } from '../../../src/xlsx.js'
import { partsFile } from '../parts.js'
import { bookParts, made, type PartsSet } from './cases.js'

const set = process.argv[2]
const built = (await setModule<PartsSet>('parts', set)).buildPartsCases()
const { books, cases, threads } = built
const out: Record<string, unknown> = {}
for (const c of cases) {
  const read = c.raw ? await readXlsx(await partsFile(c, built)) : undefined
  const book = made(read?.book ?? books[c.book], c.make, '?', read ? undefined : threads?.[c.book])
  const r = c.op ? editStructure(book, c.op) : { ok: true as const, book }
  const issues = read?.issues.map((i) => `${i.code} ${i.sheet ?? ''} ${(i.cells ?? []).join(' ')}: ${i.message}`)
  out[c.id] = r.ok ? { ...Object.fromEntries(r.book.sheets.map((s) => [s.name, bookParts(s)])), ...(issues ? { issues } : {}) } : { refused: r.code }
}
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), out)
console.log(`${set}: ${cases.length} cases predicted`)
