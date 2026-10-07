/**
 * The xlsx subpath against Excel, without Excel: the files Excel saved (`npx tsx bench/coverage/run.ts xlsx …`,
 * stored in `test/excel/xlsx/`) are read with the reader. Round trip: each book the writer wrote was opened by Excel
 * without its repair window, and what Excel saved comes back as the book written, with Excel's values equal to the
 * engine's. Case books: the books of the case sets as Excel saved them read into books that compute every case as
 * the books sent to Excel do, with the values Excel stored equal to its stored answers. Colours: the reader's fill
 * colours are the ones Excel showed.
 */
import { createHash } from 'node:crypto'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { colorBook, colorCases } from '../bench/coverage/xlsx/colors.js'
import { buildRoundTrips, XLSX_CLOCK } from '../bench/coverage/xlsx/books.js'
import { compareBooks, compareCached, entries, type Difference } from '../bench/coverage/xlsx/compare.js'
import { PART, sentBook } from '../bench/coverage/xlsx/saved.js'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from '../bench/coverage/probes.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK, matches, show, toVal } from './excel-cases.js'

const dir = new URL('./excel/xlsx/', import.meta.url)
const json = <T>(name: string): T => JSON.parse(readFileSync(new URL(name, dir), 'utf8')) as T

/** Where what Excel saved differs from the book written, and why. */
const KNOWN_BOOK: Record<string, { why: string; places: string[] }> = {
  values: { why: 'Excel writes the literal characters of a number format escaped (dd\\.mm\\.yyyy): the same format', places: ['cell Values!E3', 'cell Values!E4', 'cell Values!E5', 'cell Values!E6'] },
  'struct-edge-edge': { why: 'Excel writes a range over whole columns or rows as one (A1:A1048576 as A:A, A:XFD as $1:$1048576): the same cells', places: ['cell F!A3', 'cell F!A5', 'cell F!A15'] },
}

/** The issues of reading what Excel saved, by book: what the book itself makes the reader report. */
const ISSUES: Record<string, string[]> = {
  'formulas-values': ['errorConstant Data C2 C12 C23'],
  'spills-values': ['errorValue Spill A6', 'errorConstant Spill I1 N1'],
  styles: ['hidden Styles 9:9 F:F'],
}

describe.skipIf(!existsSync(new URL('roundtrip.json', dir)))('xlsx round trip through Excel', async () => {
  const run = json<StoredRun>('roundtrip.json')
  for (const c of await buildRoundTrips()) {
    it(`${c.id}: ${c.what}`, async () => {
      const result = compute(c.book, { clock: XLSX_CLOCK })
      const { data } = await writeXlsx(c.book, { result, mode: c.mode })
      const sent = run.books[c.id]
      expect(sent, 'no Excel answer: run the xlsx roundtrip step').toBeDefined()
      expect(sent.sent, 'the writer changed since Excel saw the file: run the xlsx roundtrip step').toBe(createHash('sha256').update(data).digest('hex'))
      expect(sent.rejected).toBeUndefined()
      const read = await readXlsx(readFileSync(new URL(`roundtrip/${c.id}.xlsx`, dir)))
      const differences: Difference[] = compareBooks(entries(c.book, c.mode === 'values' ? result : undefined), entries(read.book))
      expect(differences.map((d) => d.place)).toEqual(KNOWN_BOOK[c.id]?.places ?? [])
      // a volatile book's random numbers are Excel's own: they are compared as numbers
      const random = (d: Difference) => c.volatile && !Number.isNaN(Number(d.written)) && !Number.isNaN(Number(d.read))
      if (c.mode !== 'values') expect(compareCached(result, read.cached).filter((d) => !random(d))).toEqual([])
      expect(read.issues.map((i) => [i.code, i.sheet ?? i.name, ...(i.cells ?? [])].join(' '))).toEqual(ISSUES[c.id] ?? [])
    })
  }
})

describe.skipIf(!existsSync(new URL('colors.json', dir)))('xlsx fill colours as Excel shows them', () => {
  const stored = json<{ colors: Record<string, string> }>('colors.json').colors
  const channels = (hex: string) => [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16))

  it('palette colours and untinted theme colours exactly, tinted ones within 3 of 255 per channel', async () => {
    const cases = colorCases()
    const r = await readXlsx(await colorBook(cases))
    const off: string[] = []
    let exact = 0
    for (const c of cases) {
      const cell = r.book.sheets[0].cells[c.cell]
      const got = cell?.style === undefined ? 'none' : r.book.styles![cell.style].fill!.color
      const want = stored[c.color]
      if (got === want) exact++
      else if (!c.color.includes('tint') || channels(got).some((x, i) => Math.abs(x - channels(want)[i]) > 3)) off.push(`${c.color}: Excel ${want}, reader ${got}`)
    }
    expect(off).toEqual([])
    expect(exact).toBeGreaterThanOrEqual(212)
  })
})

/** Cases whose value differs between the book Excel saved and the book sent, and why: Excel's doing, not the reader's. */
const SAVED_KNOWN: Record<string, string> = {
  'let-held:held#5':
    "Excel saves a LET that ends with a value with its last name as the calculation (LET(x,1) as LET(x,1,x)): with 127 names that makes 255 arguments, one past the engine's limit, and the formula read does not parse",
}

/** Cases where the value Excel stored on saving differs from its stored answer: Excel does not repeat it. */
const SAVED_UNSTABLE: Record<string, string> = { 'fuzz:1#1263': 'DAY(WEEKNUM(40977,1E+3)<>2.5): 1 or #NUM! from run to run (bench/coverage/README.md)' }

describe.skipIf(!existsSync(new URL('saved.json', dir)))('xlsx case books saved by Excel', async () => {
  const files = readdirSync(new URL('saved/', dir)).filter((f) => f.endsWith('.xlsx'))
  const sets = new Map<string, { probes: Probe[]; sheets: SheetCells; answers: Record<string, Parameters<typeof matches>[1]> }>()
  const set = async (name: string) => {
    if (!sets.has(name)) {
      const mod = name === 'probes' ? { buildProbes, SHEETS: PROBE_SHEETS } : ((await import(`../bench/coverage/sets/${name}.ts`)) as { buildProbes: () => Probe[]; SHEETS?: SheetCells })
      const answers = json<{ values: Record<string, Parameters<typeof matches>[1]> }>(name === 'probes' ? '../probes.json' : `../cases/${name}.json`).values
      sets.set(name, { probes: mod.buildProbes(), sheets: mod.SHEETS ?? {}, answers })
    }
    return sets.get(name)!
  }

  it('every part of every set with stored answers went through Excel', () => {
    const names = ['probes', ...readdirSync(new URL('../cases/', dir)).filter((f) => f.endsWith('.json')).map((f) => f.slice(0, -5))]
    expect(names.filter((n) => !files.some((f) => f.startsWith(`${n}-`)))).toEqual([])
  })

  for (const file of files.sort()) {
    const [, name, k] = /^(.*)-(\d+)\.xlsx$/.exec(file)!
    it(file, async () => {
      const s = await set(name)
      const probes = s.probes.slice((Number(k) - 1) * PART, Number(k) * PART)
      const read = await readXlsx(readFileSync(new URL(`saved/${file}`, dir)))
      const c = read.book.sheets.findIndex((x) => x.name === 'C')
      const got = compute(read.book, { clock: CLOCK }).sheets[c].cells
      const sent = compute(sentBook(probes, s.sheets), { clock: CLOCK }).sheets.at(-1)!.cells
      const off: string[] = []
      probes.forEach((p, i) => {
        const key = `${name}:${p.id}`
        const [engine, fromFile] = [toVal(sent[`A${i + 1}`].value), toVal(got[`A${i + 1}`].value)]
        if (show(engine) !== show(fromFile) && !SAVED_KNOWN[key]) off.push(`${key} ${p.formula}: sent ${show(engine)}, read ${show(fromFile)}`)
        const stored = read.cached[c].cells[`A${i + 1}`]
        if ((stored === undefined || !matches(p.fn, s.answers[p.id], toVal(stored))) && !SAVED_UNSTABLE[key]) off.push(`${key}: Excel's answer ${show(s.answers[p.id])}, stored in the file ${stored === undefined ? 'nothing' : show(toVal(stored))}`)
      })
      expect(off).toEqual([])
    })
  }
})
