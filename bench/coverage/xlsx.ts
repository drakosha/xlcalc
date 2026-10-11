/**
 * The xlsx subpath against Excel. `npx tsx bench/coverage/run.ts xlsx roundtrip`: the writer writes each book of
 * `xlsx/books.ts` (with its values computed by the engine), Excel opens the file, computes it and saves it, and
 * the saved file goes to `test/excel/xlsx/roundtrip/<id>.xlsx`; a book Excel refuses (its repair window) is
 * recorded as rejected. `xlsx saved`: the books of the case sets (`sets/*.ts` with stored answers, the probes),
 * as the harness writes them for Excel, go through Excel the same way into `test/excel/xlsx/saved/`. The test
 * `test/xlsx-excel.test.ts` reads the saved files with the reader and compares them without Excel. `xlsx colors`:
 * the fill colours Excel shows for theme colours with tints and for the legacy palette (`xlsx/colors.ts`) into
 * `test/excel/xlsx/colors.json`. `xlsx formats`: the codes Excel en-US shows for the built-in number formats
 * (`xlsx/formats.ts`) into `test/excel/xlsx/formats.json`. `xlsx intersect`: formulas of the old kind computed and saved by Excel
 * (`xlsx/intersect.ts`) into `test/excel/xlsx/intersect/`. `xlsx filters`: files with filters in each form a file holds
 * them (`xlsx/filters.ts`) computed and saved by Excel into `test/excel/xlsx/filters/`.
 * `test/excel/xlsx/<step>.json` records what was sent (a hash per book), so a run sends again only the books that
 * changed (`FCOV_FRESH=1` sends all); `FXLSX_ONLY=id,id` limits a run.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { compute } from '../../src/index.js'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { inParts, writeBook } from './book.js'
import { ensureExcel, excelVersion } from './excel.js'
import { localeMeta } from './locales.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from './probes.js'
import { buildRoundTrips, XLSX_CLOCK } from './xlsx/books.js'
import { colorBook, colorCases } from './xlsx/colors.js'
import { compareBooks, compareCached, entries } from './xlsx/compare.js'
import { fillColors, numberFormats, saveThroughExcel } from './xlsx/excel.js'
import { FILTER_FILES, filterFile } from './xlsx/filters.js'
import { FORMAT_IDS, formatBook } from './xlsx/formats.js'
import { OLD_FORMULAS, oldFormulaBook } from './xlsx/intersect.js'
import { partsHash } from './xlsx/sent.js'
import { loadStored, secondsSince, setModule, writeAnswers } from './stored.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const DATA = join(ROOT, 'test/excel/xlsx')
const OUT = join(HERE, 'out/xlsx')
const ONLY = process.env.FXLSX_ONLY?.split(',').filter(Boolean)

export interface Sent {
  /** sha256 of what was sent: the written file, or a part's case formulas. */
  sent: string
  /** The repair window's text when Excel refused the file. */
  rejected?: string
}

export interface StoredRun {
  meta: Record<string, string | number>
  books: Record<string, Sent>
}

const sha = (x: Uint8Array | string) => createHash('sha256').update(x).digest('hex')

const load = (step: string): StoredRun => loadStored<StoredRun>(join(DATA, `${step}.json`)) ?? { meta: {}, books: {} }

/** Sends the books that changed through Excel in each book's locale (a tag), storing what Excel saved under `dir`. */
async function send(step: string, books: { id: string; locale: string; bytes: () => Promise<Uint8Array>; hash: string }[], report: (id: string, saved: Uint8Array) => Promise<string>) {
  const run = load(step)
  const dir = join(DATA, step)
  mkdirSync(dir, { recursive: true })
  mkdirSync(OUT, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(join(DATA, `${step}.json`), run)
  }
  // books the step no longer sends (a set that shrank) go with their files
  for (const id of Object.keys(run.books))
    if (!books.some((b) => b.id === id)) {
      delete run.books[id]
      rmSync(join(dir, `${id}.xlsx`), { force: true })
    }
  for (const locale of [...new Set(books.map((b) => b.locale))].sort()) {
    const todo = books.filter((b) => b.locale === locale && (!ONLY || ONLY.includes(b.id)) && !(run.books[b.id]?.sent === b.hash && (run.books[b.id].rejected || existsSync(join(dir, `${b.id}.xlsx`)))))
    if (!todo.length) continue
    await ensureExcel(locale)
    for (const b of todo) {
      const bytes = await b.bytes()
      writeFileSync(join(OUT, `${step}-${b.id}.xlsx`), bytes)
      const r = await saveThroughExcel(bytes, b.id)
      if ('rejected' in r) {
        run.books[b.id] = { sent: b.hash, rejected: r.rejected }
        console.log(`${b.id}: Excel refused the file: ${r.rejected}`)
      } else {
        writeFileSync(join(dir, `${b.id}.xlsx`), r.saved)
        run.books[b.id] = { sent: b.hash }
        console.log(`${b.id}: saved, ${await report(b.id, r.saved)}`)
      }
      save()
    }
  }
  save()
}

async function roundTrip() {
  const cases = await buildRoundTrips()
  const books = await Promise.all(
    cases.map(async (c) => {
      const result = compute(c.book, { clock: XLSX_CLOCK })
      const { data } = await writeXlsx({ book: c.book, result }, { mode: c.mode })
      return { id: c.id, locale: c.book.locale ?? 'en-US', bytes: async () => data, hash: await partsHash(data), c, result }
    }),
  )
  await send('roundtrip', books, async (id, saved) => {
    const b = books.find((x) => x.id === id)!
    const read = await readXlsx(saved)
    const book = compareBooks(entries(b.c.book, b.c.mode === 'values' ? b.result : undefined), entries(read.book))
    const values = b.c.mode === 'values' ? [] : compareCached(b.result, read.cached)
    return `${book.length} book differences, ${values.length} value differences, ${read.issues.length} issues`
  })
}

/** The case sets with stored en-US answers, and the probes, in parts of 400 cases as the harness sends them. */
async function caseParts(): Promise<{ id: string; probes: Probe[]; sheets: SheetCells }[]> {
  const names = readdirSync(join(ROOT, 'test/excel/cases'))
    .filter((f) => f.endsWith('.json'))
    .map((f) => f.slice(0, -5))
    .sort()
  const sets: { name: string; probes: Probe[]; sheets: SheetCells }[] = [{ name: 'probes', probes: buildProbes(), sheets: PROBE_SHEETS }]
  for (const name of names) {
    const mod = await setModule<{ buildProbes: () => Probe[]; SHEETS?: SheetCells }>('sets', name)
    sets.push({ name, probes: mod.buildProbes(), sheets: mod.SHEETS ?? {} })
  }
  return sets.flatMap((s) => inParts(s.probes).map((probes, k) => ({ id: `${s.name}-${k + 1}`, probes, sheets: s.sheets })))
}

async function saved() {
  const parts = await caseParts()
  const books = parts.map((p) => ({
    id: p.id,
    locale: 'en-US',
    hash: sha(JSON.stringify({ formulas: p.probes.map((x) => x.formula), sheets: p.sheets })),
    bytes: async () => {
      const file = join(OUT, `case-${p.id}.xlsx`)
      await writeBook(p.probes, file, p.sheets)
      return new Uint8Array(readFileSync(file))
    },
  }))
  await send('saved', books, async (_, file) => `${(await readXlsx(file)).issues.length} issues`)
}

async function colors() {
  const cases = colorCases()
  await ensureExcel('en-US')
  const shown = await fillColors(await colorBook(cases), cases.map((c) => c.cell))
  const meta = { version: excelVersion(), locale: localeMeta('en-US'), date: new Date().toISOString().slice(0, 10) }
  writeAnswers(join(DATA, 'colors.json'), { meta, colors: Object.fromEntries(cases.map((c, i) => [c.color, shown[i]])) })
  console.log(`colors: ${cases.length} fills read`)
}

async function formats() {
  await ensureExcel('en-US')
  const shown = await numberFormats(await formatBook(), FORMAT_IDS.map((id) => `A${id}`))
  const meta = { version: excelVersion(), locale: localeMeta('en-US'), date: new Date().toISOString().slice(0, 10) }
  writeAnswers(join(DATA, 'formats.json'), { meta, formats: Object.fromEntries(FORMAT_IDS.map((id, i) => [id, shown[i]])) })
  console.log(`formats: ${shown.length} read`)
}

async function intersect() {
  const data = await oldFormulaBook(OLD_FORMULAS)
  await send('intersect', [{ id: 'old-formulas', locale: 'en-US', bytes: async () => data, hash: await partsHash(data) }], async (_, saved) => `${(await readXlsx(saved)).issues.length} issues`)
}

async function filters() {
  const books = await Promise.all(FILTER_FILES.map(async (c) => ({ id: c.id, locale: 'en-US', data: await filterFile(c) })))
  await send(
    'filters',
    await Promise.all(books.map(async (b) => ({ ...b, bytes: async () => b.data, hash: await partsHash(b.data) }))),
    async (_, saved) => {
      const read = await readXlsx(saved)
      const values = compareCached(compute(read.book, { clock: XLSX_CLOCK }), read.cached)
      return `${values.length} value differences, ${read.book.sheets[0].filtered ? 'filtered' : 'not filtered'}`
    },
  )
}

export async function stepXlsx(what: string | undefined): Promise<void> {
  if (what === 'roundtrip') await roundTrip()
  else if (what === 'filters') await filters()
  else if (what === 'saved') await saved()
  else if (what === 'colors') await colors()
  else if (what === 'intersect') await intersect()
  else if (what === 'formats') await formats()
  else throw new Error('xlsx: roundtrip | saved | filters | colors | formats | intersect')
}
