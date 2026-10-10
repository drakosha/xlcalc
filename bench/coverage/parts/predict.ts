/**
 * The engine's answers to a parts set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/parts/predict.ts <set>` writes `parts/<set>.predicted.json`, for each case the engine's refusal
 * or each sheet's notes, hyperlinks and frozen panes after the operation (what Excel makes is given as made, with its
 * author unknown: `?`).
 */
import { writeFileSync } from 'node:fs'
import { restructure } from '../../../src/index.js'
import type { BookInput } from '../../../src/contract.js'
import { bookParts, type PartsMake, type PartsSet, type ThreadSpec } from './cases.js'

/**
 * The book with what Excel makes and the threads of its file, as the engine holds them: a thread as a note, its first
 * author's, its comments' text with the replies after an empty line with their authors.
 */
export function made(book: BookInput, make: PartsMake[] = [], author = '?', threads: ThreadSpec[] = []): BookInput {
  const sheets = book.sheets.map((s) => ({ ...s, cells: { ...s.cells } }))
  const sheet = (name: string) => sheets.find((s) => s.name === name)!
  for (const m of make) {
    if ('note' in m) sheet(m.sheet).cells[m.note] = { ...sheet(m.sheet).cells[m.note], note: { author, text: m.text } }
    else if ('link' in m) {
      const { link: cell, sheet: _, ...link } = m
      sheet(m.sheet).cells[cell] = { ...sheet(m.sheet).cells[cell], link }
    } else sheet(m.freeze).frozen = { ...(m.rows ? { rows: m.rows } : {}), ...(m.columns ? { columns: m.columns } : {}) }
  }
  for (const t of threads) {
    const text = t.comments.map((c, i) => (i ? `${c.author}: ${c.text}` : c.text)).join('\n\n')
    sheet(t.sheet).cells[t.cell] = { ...sheet(t.sheet).cells[t.cell], note: { author: t.comments[0].author, text } }
  }
  return { ...book, sheets }
}

if (process.argv[1] === new URL(import.meta.url).pathname) {
  const set = process.argv[2]
  if (!/^[a-z0-9-]+$/.test(set ?? '')) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const { books, cases, threads } = ((await import(`./${set}.js`)) as PartsSet).buildPartsCases()
  const out: Record<string, unknown> = {}
  for (const c of cases) {
    const book = made(books[c.book], c.make, '?', threads?.[c.book])
    const r = c.op ? restructure(book, c.op) : { ok: true as const, book }
    out[c.id] = r.ok ? Object.fromEntries(r.book.sheets.map((s) => [s.name, bookParts(s)])) : { refused: r.code }
  }
  writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(`${set}: ${cases.length} cases predicted`)
}
