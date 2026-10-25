/**
 * The xlsx subpath against Excel, without Excel: the files Excel saved (`npx tsx bench/coverage/run.ts xlsx …`,
 * stored in `test/excel/xlsx/`) are read with the reader. Round trip: each book the writer wrote was opened by Excel
 * without its repair window, and what Excel saved comes back as the book written, with Excel's values equal to the
 * engine's. Case books: the books of the case sets as Excel saved them read into books that compute every case as
 * the books sent to Excel do, with the values Excel stored equal to its stored answers. Formulas of the old kind: read
 * with `@` where Excel 365 puts it, they compute the values Excel stored. Filters in each form a file holds them: read
 * with the filter mode Excel took, they compute the values Excel stored. The answers predicted before Excel ran
 * (`bench/coverage/xlsx/predicted.json`) are Excel's, the misses named. Colours and built-in number formats: the
 * reader's are the ones Excel showed.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { colorBook, colorCases } from '../bench/coverage/xlsx/colors.js'
import { buildRoundTrips, XLSX_CLOCK } from '../bench/coverage/xlsx/books.js'
import { compareBooks, compareCached, entries, type Difference } from '../bench/coverage/xlsx/compare.js'
import { FILTER_FILES, filterFile } from '../bench/coverage/xlsx/filters.js'
import type { Predicted } from '../bench/coverage/xlsx/predict.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { formulaCell, OLD_FORMULAS } from '../bench/coverage/xlsx/intersect.js'
import { PART, sentBook } from '../bench/coverage/xlsx/saved.js'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { buildProbes, PROBE_SHEETS, type Probe, type SheetCells } from '../bench/coverage/probes.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { BUILTIN } from '../src/xlsx/styles/numfmt.js'
import { CLOCK, matches, show, toVal } from './excel-cases.js'
import { cellsOf } from './looks-book.js'
import './locales.js'
import { filterRun } from './xlsx-filters.js'

const dir = new URL('./excel/xlsx/', import.meta.url)
const json = <T>(name: string): T => JSON.parse(readFileSync(new URL(name, dir), 'utf8')) as T

/** The built-in formats Excel for Mac saves otherwise: money and accounting (5 to 8, 41 to 44) as codes, 23 to 26 as General, 27 to 36 as 14 and 21. */
const RESAVED = [5, 6, 7, 8, ...Array.from({ length: 14 }, (_, i) => 23 + i), 41, 42, 43, 44]
const BUILTIN_RESAVED = {
  why: 'Excel for Mac saves the built-in money and accounting formats (5 to 8, 41 to 44) as codes of their own in the locale it runs in, 23 to 26 as General, and the East Asian 27 to 36 as 14 and 21',
  places: [...RESAVED.map((id) => `cell Formats!A${id}`), 'column Formats!C'],
}

/** Where what Excel saved differs from the book written, and why. */
const KNOWN_BOOK: Record<string, { why: string; places: string[] }> = {
  values: { why: 'Excel writes the literal characters of a number format escaped (dd\\.mm\\.yyyy): the same format', places: ['cell Values!E3', 'cell Values!E4', 'cell Values!E5', 'cell Values!E6'] },
  date1904: { why: 'Excel writes the literal characters of a number format escaped (yyyy\\-mm\\-dd): the same format', places: ['cell Dates!A1', 'cell Dates!A4'] },
  builtin: BUILTIN_RESAVED,
  'builtin-ru': BUILTIN_RESAVED,
  'struct-edge-edge': { why: 'Excel writes a range over whole columns or rows as one (A1:A1048576 as A:A, A:XFD as $1:$1048576): the same cells', places: ['cell F!A3', 'cell F!A5', 'cell F!A15'] },
  'spill-names-formulas': { why: 'a name that reads as a cell address (nm2) is not a name in Excel: the writer leaves it out', places: ['name !NM2'] },
  'lost-sheet': { why: 'the engine keeps a reference to a deleted sheet as Excel shows it (#REF!A1); a file holds #REF!, as Excel writes it: the same value', places: ['cell Plain!A1', 'cell Plain!A2', 'cell Plain!A3', 'name !TOTAL'] },
  'spill-names-sheet-global': { why: "Excel saves a global name qualified by a sheet without a local one as the book's own name ([0]!x), read as x: the same name", places: ['cell S!E1'] },
}

/** The issues of reading what Excel saved, by book: what the book itself makes the reader report. */
const ISSUES: Record<string, string[]> = {
  'formulas-values': ['errorConstant Data C2 C12 C23'],
  'spills-values': ['errorValue Spill A6', 'errorConstant Spill I1 N1'],
  filtered: ['autoFilter S'],
  'filtered-table': ['autoFilter S'],
}

const predicted = JSON.parse(readFileSync(new URL('../bench/coverage/xlsx/predicted.json', import.meta.url), 'utf8')) as Predicted

describe.skipIf(!existsSync(new URL('roundtrip.json', dir)))('xlsx round trip through Excel', async () => {
  const run = json<StoredRun>('roundtrip.json')
  for (const c of await buildRoundTrips()) {
    it(`${c.id}: ${c.what}`, async () => {
      const result = compute(c.book, { clock: XLSX_CLOCK })
      const { data } = await writeXlsx({ book: c.book, result }, { mode: c.mode })
      const sent = run.books[c.id]
      expect(sent, 'no Excel answer: run the xlsx roundtrip step').toBeDefined()
      expect(sent.sent, 'the writer changed since Excel saw the file: run the xlsx roundtrip step').toBe(await partsHash(data))
      expect(sent.rejected).toBeUndefined()
      const read = await readXlsx(readFileSync(new URL(`roundtrip/${c.id}.xlsx`, dir)))
      const differences: Difference[] = compareBooks(entries(c.book, c.mode === 'values' ? result : undefined), entries(read.book))
      expect(differences.map((d) => d.place)).toEqual(KNOWN_BOOK[c.id]?.places ?? [])
      // a volatile book's random numbers are Excel's own: they are compared as numbers
      const random = (d: Difference) => c.volatile && !Number.isNaN(Number(d.written)) && !Number.isNaN(Number(d.read))
      if (c.mode !== 'values') expect(compareCached(result, read.cached).filter((d) => !random(d))).toEqual([])
      expect(read.issues.map((i) => [i.code, i.sheet ?? i.name, ...(i.cells ?? [])].join(' '))).toEqual(ISSUES[c.id] ?? [])
      // a book predicted before Excel ran: Excel stored the values predicted; the book came back as predicted, but where Excel saves otherwise
      const p = predicted.roundtrip[c.id]
      if (p) {
        expect(Object.fromEntries(Object.keys(p.values).map((k) => [k, read.cached.sheets[read.book.sheets.findIndex((x) => x.name === k.split('!')[0])].cells[k.split('!')[1]]?.value]))).toEqual(p.values)
        expect(differences.map((d) => d.place).filter((x) => !KNOWN_BOOK[c.id]?.places.includes(x))).toEqual(p.differences)
      }
    })
  }
})

describe.skipIf(!existsSync(new URL('filters.json', dir)))('xlsx filters as Excel computed and saved them', () => {
  const run = json<StoredRun>('filters.json')

  filterRun(run, 'Excel', async (c) => {
    const read = await readXlsx(readFileSync(new URL(`filters/${c.id}.xlsx`, dir)))
    // the filter mode Excel took, as predicted: the engine computes the values Excel stored, which are the ones predicted
    expect(!!read.book.sheets[0].filtered).toBe(predicted.filters[c.id].filtered)
    expect(compareCached(compute(read.book, { clock: XLSX_CLOCK }), read.cached)).toEqual([])
    expect(Object.fromEntries(Object.entries(read.cached.sheets[1].cells).map(([a, v]) => [`F!${a}`, v.value]))).toEqual(predicted.filters[c.id].values)
  })
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
      const cell = cellsOf(r.book.sheets[0])[c.cell]
      const got = cell?.style === undefined ? 'none' : r.book.styles![cell.style].fill!.color
      const want = stored[c.color]
      if (got === want) exact++
      else if (!c.color.includes('tint') || channels(got).some((x, i) => Math.abs(x - channels(want)[i]) > 3)) off.push(`${c.color}: Excel ${want}, reader ${got}`)
    }
    expect(off).toEqual([])
    expect(exact).toBeGreaterThanOrEqual(212)
  })
})

/** Excel for Mac's reading of the long s, a known difference. */
const MAC_LONG_S = 'Excel for Mac reads the long s as s and saves falſe as FALSE, which the formula read is; the engine, as Excel for Windows, takes falſe for a name (docs/limitations.md, Known differences)'

/** Cases whose value differs between the book Excel saved and the book sent, and why: Excel's doing, not the reader's. */
const SAVED_KNOWN: Record<string, string> = {
  'let-held:held#5':
    "Excel saves a LET that ends with a value with its last name as the calculation (LET(x,1) as LET(x,1,x)): with 127 names that makes 255 arguments, one past the engine's limit, and the formula read does not parse",
  'name-case-held:bool#1': MAC_LONG_S,
  'name-case-held:bool#2': MAC_LONG_S,
  'name-case-held:bool#3': MAC_LONG_S,
  'name-case-held:bool#4': MAC_LONG_S,
  'name-case-held:bool#5': MAC_LONG_S,
  'name-case-held:bool#6': MAC_LONG_S,
  'name-case-held:let#2': MAC_LONG_S,
}

/** Cases where the value Excel stored on saving differs from its stored answer: Excel does not repeat it. */
const SAVED_UNSTABLE: Record<string, string> = { 'fuzz:1#1263': 'DAY(WEEKNUM(40977,1E+3)<>2.5): 1 or #NUM! from run to run (bench/coverage/README.md)' }

/** Cells Excel left uncomputed, in a circular reference without iterative calculation: the file stores empty text for them. */
const SAVED_UNCOMPUTED: Record<string, string> = Object.fromEntries(['SUMIF#1', 'SUMIF#2', 'SUMIF#3'].map((id) => [`sumif-own-held:${id}`, "SUMIF inside its own resized sum range (OPEN in test/excel-open.ts)"]))

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

  it('every part of every set with stored answers went through Excel, and only those', async () => {
    const names = ['probes', ...readdirSync(new URL('../cases/', dir)).filter((f) => f.endsWith('.json')).map((f) => f.slice(0, -5))]
    const parts: string[] = []
    for (const name of names) for (let k = 0; k * PART < (await set(name)).probes.length; k++) parts.push(`${name}-${k + 1}.xlsx`)
    expect([...files].sort()).toEqual(parts.sort())
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
        const stored = read.cached.sheets[c].cells[`A${i + 1}`]?.value
        if ((stored === undefined || !matches(p.fn, s.answers[p.id], toVal(stored))) && !SAVED_UNSTABLE[key] && !SAVED_UNCOMPUTED[key]) off.push(`${key}: Excel's answer ${show(s.answers[p.id])}, stored in the file ${stored === undefined ? 'nothing' : show(toVal(stored))}`)
      })
      expect(off).toEqual([])
    })
  }
})

describe.skipIf(!existsSync(new URL('intersect.json', dir)))('xlsx formulas of the old kind against Excel', () => {
  it('every formula of the old kind read computes the value Excel stored, but those of functions the engine does not have and those it does not parse', async () => {
    const read = await readXlsx(readFileSync(new URL('intersect/old-formulas.xlsx', dir)))
    const unread = new Set(read.issues.filter((i) => i.code === 'formula').flatMap((i) => i.cells ?? []))
    const engine = compute(read.book, { clock: CLOCK }).sheets[0].cells
    const off = OLD_FORMULAS.flatMap((f, i) => {
      const a = formulaCell(i)
      const [mine, excel] = [toVal(engine[a].value), toVal(read.cached.sheets[0].cells[a]?.value ?? null)]
      return show(mine) !== show(excel) && !unread.has(a) && !(mine.t === 'e' && mine.v === 'NAME') ? [`${a} ${f} read as ${read.book.sheets[0].cells[a].formula}: engine ${show(mine)}, Excel ${show(excel)}`] : []
    })
    expect(off).toEqual([])
    expect([...unread]).toEqual([])
  })
})

describe.skipIf(!existsSync(new URL('formats.json', dir)))('xlsx built-in number formats as Excel shows them', () => {
  it('the reader reads each built-in format as the code Excel en-US gives it: Windows\'s to 49, the Mac\'s from 50', () => {
    const stored = json<{ formats: Record<string, string> }>('formats.json').formats
    // what Excel for Windows reports for the formats 0 to 49 (its short date is m/d/yyyy, the Mac's m/d/yy)
    const win = JSON.parse(readFileSync(new URL('../win/display/builtin.json', dir), 'utf8')) as { cases: { id: string; numFmtId: number }[]; answers: Record<string, { code: string }> }
    const windows = Object.fromEntries(win.cases.map((c) => [String(c.numFmtId), win.answers[c.id].code]))
    expect({ ...BUILTIN }).toEqual(Object.fromEntries(Object.entries(stored).map(([id, code]) => [id, Number(id) <= 49 ? windows[id] : code])))
  })
})
