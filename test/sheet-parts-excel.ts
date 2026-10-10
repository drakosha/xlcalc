/**
 * Excel's answers to notes, hyperlinks and frozen panes (`test/excel/parts/<set>.json` and the files Excel saved beside
 * them, from `run.ts parts <set>`; Excel for Windows's in `test/excel/win/parts/`) and their comparison with a book's
 * parts: notes by cell with author and text, hyperlinks by cell with address, place and tooltip, frozen panes by the rows
 * and columns before their split.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { partsFile, type StoredParts } from '../bench/coverage/parts.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { bookParts, type PartsAnswer, type PartsMake, type PartsSheetAnswer, type SheetParts } from '../bench/coverage/parts/cases.js'
import type { PartsCase, PartsCases, PartsSet } from '../bench/coverage/parts/cases.js'
import { made } from '../bench/coverage/parts/predict.js'
import { restructure } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import type { BookInput, CellLink, FrozenPanes } from '../src/contract.js'
import { formatAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'

export type PartsSetFile = StoredParts & { name: string; dir: URL }

/** The stored sets of a directory (`excel/parts/`, `excel/win/parts/`) with the folders of their saved files. */
export function partsSets(sub: string): PartsSetFile[] {
  const dir = new URL(`./excel/${sub}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .map((f) => ({ name: f.slice(0, -5), dir: new URL(`${f.slice(0, -5)}/`, dir), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredParts) }))
}

/** The file Excel saved for a case. */
export const savedFile = (set: PartsSetFile, id: string): Uint8Array => new Uint8Array(readFileSync(new URL(`${id}.xlsx`, set.dir)))

/** A thread of comments as the reader takes it: the first author, the comments' text, the replies after an empty line with their authors. */
const threadNote = (t: { author: string; text: string }[]) => ({ ...(t[0].author ? { author: t[0].author } : {}), text: t.map((c, i) => (i ? `${c.author}: ${c.text}` : c.text)).join('\n\n') })

/** What Excel read on a sheet, as a book's parts: a thread of comments as a note, a hyperlink on each of its cells. */
export function excelParts(s: PartsSheetAnswer): SheetParts {
  const out: SheetParts = { notes: {}, links: {} }
  for (const [addr, n] of Object.entries(s.notes)) out.notes[addr] = n.author ? { author: n.author, text: n.text } : { text: n.text }
  for (const [addr, t] of Object.entries(s.threads ?? {})) out.notes[addr] = threadNote(t)
  for (const l of s.links) {
    const link: CellLink = { ...(l.address ? { url: l.address } : {}), ...(l.subAddress ? { location: l.subAddress } : {}), ...(l.tip ? { tooltip: l.tip } : {}) }
    const r = parseRange(l.range)!
    for (let row = r.r1; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) out.links[formatAddress(row, c)] = link
  }
  const p = s.panes
  if (p.freeze && (p.splitRow || p.splitColumn)) {
    const frozen: FrozenPanes = {}
    if (p.splitRow) frozen.rows = p.scrollRow - 1 + p.splitRow
    if (p.splitColumn) frozen.columns = p.scrollColumn - 1 + p.splitColumn
    out.frozen = frozen
  }
  return out
}

/** The author Excel gives the notes it makes itself: the author of the first one it made, or `?`. */
export function madeAuthor(make: PartsMake[] | undefined, a: PartsAnswer): string {
  for (const m of make ?? []) if ('note' in m) return a.sheets.find((x) => x.name === m.sheet)?.notes[m.note]?.author ?? '?'
  return '?'
}

/** The book of a case with what Excel made (the author Excel's) and the threads of its file, as the engine holds them. */
export const madeBook = (set: PartsCases, c: PartsCase, a: PartsAnswer): BookInput => made(set.books[c.book], c.make, madeAuthor(c.make, a), set.threads?.[c.book])

const show = (v: unknown) => (v === undefined ? 'none' : JSON.stringify(v))

/** The lines where Excel's sheets and a book's parts differ: a sheet missing on either side, a note, a link, the frozen panes. */
export function partsDifferences(sheets: { name: string; parts: SheetParts }[], book: { name: string; parts: SheetParts }[]): string[] {
  const out: string[] = []
  const names = [...new Set([...sheets.map((s) => s.name), ...book.map((s) => s.name)])]
  for (const name of names) {
    const [x, y] = [sheets.find((s) => s.name === name)?.parts, book.find((s) => s.name === name)?.parts]
    if (!x || !y) {
      out.push(`${name}: ${x ? 'only in Excel' : 'only in the book'}`)
      continue
    }
    for (const kind of ['notes', 'links'] as const)
      for (const addr of [...new Set([...Object.keys(x[kind]), ...Object.keys(y[kind])])].sort()) {
        const [a, b] = [show(x[kind][addr]), show(y[kind][addr])]
        if (a !== b) out.push(`${name} ${kind === 'notes' ? 'note' : 'link'} ${addr}: Excel ${a}, book ${b}`)
      }
    const [a, b] = [show(x.frozen && { rows: x.frozen.rows, columns: x.frozen.columns }), show(y.frozen && { rows: y.frozen.rows, columns: y.frozen.columns })]
    if (a !== b) out.push(`${name} frozen: Excel ${a}, book ${b}`)
  }
  return out
}

/** A book's sheets as parts, in the comparison's shape. */
export const sheetsOf = (book: BookInput) => book.sheets.map((s) => ({ name: s.name, parts: bookParts(s) }))

/** Excel's sheets as parts. */
export const excelSheets = (a: PartsAnswer) => a.sheets.map((s) => ({ name: s.name, parts: excelParts(s) }))

/** The engine's book after a case, or its refusal. */
function engineAfter(set: PartsSetFile, id: string) {
  const c = set.cases.find((x) => x.id === id)!
  const book = madeBook(set, c, set.answers[id])
  return c.op ? restructure(book, c.op) : { ok: true as const, book }
}

/**
 * The lines about the cells where Excel showed a thread's placeholder note (author `tc={id}`), which neither Excel's
 * object model takes for a thread, set apart from the others: each such cell must differ, and only by its note.
 */
function apart(a: PartsAnswer, lines: string[]): { placeholders: string[]; rest: string[] } {
  const cells = new Set(a.sheets.flatMap((s) => Object.entries(s.notes).flatMap(([cell, n]) => (n.author.startsWith('tc=') ? [`${s.name} note ${cell}:`] : []))))
  const at = (l: string) => l.slice(0, l.indexOf(':') + 1)
  return { placeholders: lines.filter((l) => cells.has(at(l))).map(at), rest: lines.filter((l) => !cells.has(at(l))) }
}

/**
 * The tests of stored sets: each case answered, Excel read what the engine holds, the reader reads Excel's file as Excel
 * read it and reports nothing of it, and reads the threads of comments in it as notes (where the set has threads, the
 * engine's book being the reference: Excel shows each thread's placeholder note, `apart`); `known`, `readKnown` and
 * `threadsKnown` pin the differences by `<set>:<case id>`, `issues` what the reader reports (`code sheet cells`).
 */
export function checkSets(sets: PartsSetFile[], known: Readonly<Record<string, string[]>>, readKnown: Readonly<Record<string, string[]>>, from: (name: string) => Promise<PartsSet>, issues: Readonly<Record<string, string[]>> = {}, threadsKnown: Readonly<Record<string, string[]>> = {}) {
  for (const set of sets) {
    describe(`parts set ${set.name}`, () => {
      it('the stored answers are for the current cases and books', async () => {
        const built = (await from(set.name)).buildPartsCases()
        expect(set.cases).toEqual(built.cases)
        expect(set.books).toEqual(built.books)
        expect(set.threads).toEqual(built.threads)
        expect(built.cases.filter((c) => !set.answers[c.id] && !set.rejected?.[c.id])).toEqual([])
        for (const c of built.cases) expect(set.sent[c.id], `${c.id}: the writer changed since Excel saw the file`).toBe(await partsHash(await partsFile(c, built)))
      })
      for (const c of set.cases) {
        const key = `${set.name}:${c.id}`
        it(`${c.id}: Excel loaded the book`, () => expect(set.rejected?.[c.id]).toBeUndefined())
        if (!set.answers[c.id]) continue
        it(`${c.id}: Excel read what the engine holds`, () => {
          const a = set.answers[c.id]
          const r = engineAfter(set, c.id)
          if (a.error) return void expect(r.ok, `Excel refused: ${a.error}`).toBe(false)
          const lines = r.ok ? partsDifferences(excelSheets(a), sheetsOf(r.book)) : [`engine refused (${r.code}): ${r.message}`]
          const { placeholders, rest } = apart(a, lines)
          expect(rest).toEqual(known[key] ?? [])
          expect(placeholders.length, "the threads' cells, where Excel shows their placeholder notes").toBe(set.threads?.[c.book]?.length ?? 0)
        })
        if (set.threads?.[c.book])
          it(`${c.id}: the reader reads the threads of Excel's file as notes`, async () => {
            const r = engineAfter(set, c.id)
            expect(r.ok).toBe(true)
            expect(partsDifferences(sheetsOf((await readXlsx(savedFile(set, c.id))).book), r.ok ? sheetsOf(r.book) : [])).toEqual(threadsKnown[key] ?? [])
          })
        it(`${c.id}: the reader reads Excel's file as Excel read it`, async () => {
          const read = await readXlsx(savedFile(set, c.id))
          const { placeholders, rest } = apart(set.answers[c.id], partsDifferences(excelSheets(set.answers[c.id]), sheetsOf(read.book)))
          expect(rest).toEqual(readKnown[key] ?? [])
          expect(placeholders.length, "the threads' cells, where Excel shows their placeholder notes").toBe(set.threads?.[c.book]?.length ?? 0)
          expect(read.issues.map((i) => `${i.code} ${i.sheet ?? ''} ${(i.cells ?? []).join(' ')}`.trim())).toEqual(issues[key] ?? [])
        })
      }
    })
  }
}

/** Excel takes the part of an address after `#` as its place (Address and SubAddress); the book keeps the address as written. */
const anchor = (cell: string) => `S link ${cell}: Excel {"url":"https://example.com/page","location":"top"}, book {"url":"https://example.com/page#top"}`

/** `<set>:<case id>` → the difference lines pinned between Excel and the engine, each with its reason in a comment; both Excels give them. */
export const PARTS_KNOWN: Readonly<Record<string, string[]>> = {
  'files:written': [anchor('D13')],
  'files:after-insert': [anchor('D14')],
  'files:after-move': [anchor('D13')],
  'files:after-copy': [anchor('D13')],
  // the writer cuts a tooltip to Excel's 255 characters: Excel does not load a file with a longer one
  'files:tip': [`S link A1: Excel ${JSON.stringify({ url: 'https://example.com/', tooltip: 't'.repeat(255) })}, book ${JSON.stringify({ url: 'https://example.com/', tooltip: 't'.repeat(300) })}`],
}

/** Excel saves a space of an address as `%20` and shows it as a space; the reader gives the address as the file holds it. */
const space = (cell: string) => `S link ${cell}: Excel {"url":"https://example.com/a b/ж"}, book {"url":"https://example.com/a%20b/ж"}`

/** `<set>:<case id>` → the difference lines pinned between Excel and the reader of the file Excel saved. */
export const PARTS_READ_KNOWN: Readonly<Record<string, string[]>> = {
  'files:written': [space('D6')],
  'files:after-insert': [space('D7')],
  'files:after-move': [space('D6')],
  'files:after-copy': [space('D6')],
}

/** The reader reports each thread of comments it reads as a note. */
const threadIssues = (cells: string) => [`comment S ${cells}`, 'comment T A1']

/** What the reader reports of the files Excel saved, by `<set>:<case id>`. */
export const THREAD_ISSUES: Readonly<Record<string, string[]>> = { 'threads:threads': threadIssues('B2 C3'), 'threads:threads-insert-row': threadIssues('B3 C4') }
