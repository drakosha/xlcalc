/**
 * The xlsx subpath against Excel, without Excel: the files Excel saved (`npx tsx bench/coverage/run.ts xlsx …`,
 * stored in `test/excel/xlsx/`) are read with the reader. Round trip: each book the writer wrote was opened by Excel
 * without its repair window, and what Excel saved comes back as the book written, with Excel's values equal to the
 * engine's. Colours: the reader's fill colours are the ones Excel showed.
 */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { colorBook, colorCases } from '../bench/coverage/xlsx/colors.js'
import { buildRoundTrips, XLSX_CLOCK } from '../bench/coverage/xlsx/books.js'
import { compareBooks, compareCached, entries, type Difference } from '../bench/coverage/xlsx/compare.js'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'

const dir = new URL('./excel/xlsx/', import.meta.url)
const json = <T>(name: string): T => JSON.parse(readFileSync(new URL(name, dir), 'utf8')) as T

/** Where what Excel saved differs from the book written, and why. */
const KNOWN_BOOK: Record<string, { why: string; places: string[] }> = {
  values: { why: 'Excel writes the literal characters of a number format escaped (dd\\.mm\\.yyyy): the same format', places: ['cell Values!E3', 'cell Values!E4', 'cell Values!E5', 'cell Values!E6'] },
}

/** The issues of reading what Excel saved, by book: what the book itself makes the reader report. */
const ISSUES: Record<string, string[]> = {
  'formulas-values': ['errorConstant Data C2 C12 C23'],
  'spills-values': ['errorConstant Spill I1 N1', 'errorValue Spill A6'],
  styles: ['hidden Styles 9:9 F:F'],
}

describe.skipIf(!existsSync(new URL('roundtrip.json', dir)))('xlsx round trip through Excel', async () => {
  const run = json<StoredRun>('roundtrip.json')
  for (const c of buildRoundTrips()) {
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
      if (c.mode !== 'values') expect(compareCached(result, read.cached)).toEqual([])
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
