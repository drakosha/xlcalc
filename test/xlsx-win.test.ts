/**
 * The xlsx filter files as Excel for Windows computed and saved them (`test/excel/win/xlsx/filters/`, `bench/windows`):
 * read back, each has the filter mode Excel for Mac took, and the engine computes the values Excel for Windows stored,
 * as `xlsx-excel.test.ts` checks the Mac's. The built-in number formats 1 to 58 as Excel for Windows shows them in
 * each locale (`xlsx/formats.json` in en-US, `xlsx/<tag>/formats.json` in another): 1 to 49 as the display set `builtin`
 * read them, 50 to 58 pinned.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { XLSX_CLOCK } from '../bench/coverage/xlsx/books.js'
import { compareCached } from '../bench/coverage/xlsx/compare.js'
import { FILTER_FILES, filterFile } from '../bench/coverage/xlsx/filters.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { compute } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { BUILTIN } from '../src/xlsx/styles/numfmt.js'
import './locales.js'

const dir = new URL('./excel/win/xlsx/', import.meta.url)
const mac = new URL('./excel/xlsx/', import.meta.url)

describe.skipIf(!existsSync(new URL('filters.json', dir)))('xlsx filters as Excel for Windows computed and saved them', () => {
  const run = JSON.parse(readFileSync(new URL('filters.json', dir), 'utf8')) as StoredRun

  it('every file of the set went through Excel for Windows, and only those', () => {
    expect(Object.keys(run.books).sort()).toEqual(FILTER_FILES.map((c) => c.id).sort())
  })

  for (const c of FILTER_FILES)
    it(`${c.id}: ${c.about}`, async () => {
      const data = await filterFile(c)
      expect(run.books[c.id].sent, 'the file changed since Excel saw it: run the xlsx filters step').toBe(await partsHash(data))
      expect(run.books[c.id].rejected).toBeUndefined()
      const read = await readXlsx(readFileSync(new URL(`filters/${c.id}.xlsx`, dir)))
      const onMac = await readXlsx(readFileSync(new URL(`filters/${c.id}.xlsx`, mac)))
      expect(!!read.book.sheets[0].filtered).toBe(!!onMac.book.sheets[0].filtered)
      expect(compareCached(compute(read.book, { clock: XLSX_CLOCK }), read.cached)).toEqual([])
    })
})

/**
 * Known difference: Excel for Windows shows the built-in formats 50 to 58 (East Asian dates) as its short date, format 14,
 * in both locales (`m/d/yyyy`, `ДД.ММ.ГГГГ`), as Excel for Mac en-US shows them as its own 14 (`m/d/yy`); the reader
 * reads them as that Mac code, a format of its own, where it reads 14 by number and shows it by the book's locale.
 */
const EAST_ASIAN_READ = 'm/d/yy'

describe.each(TAGS.map((tag) => [tag, localeDir(tag)]))('xlsx built-in number formats as Excel for Windows %s shows them', (tag, sub) => {
  const file = new URL(`${sub}formats.json`, dir)
  it.skipIf(!existsSync(file))('1 to 49 are the codes the display set builtin read; 50 to 58 are the short date, which the reader reads as the Mac\'s code', () => {
    const run = JSON.parse(readFileSync(file, 'utf8')) as { meta: { locale: string }; formats: Record<string, string>; local: Record<string, string> }
    expect(run.meta.locale).toBe(tag)
    const display = JSON.parse(readFileSync(new URL(`../display/${sub}builtin.json`, dir), 'utf8')) as { cases: { id: string; numFmtId: number }[]; answers: Record<string, { code: string; local: string }> }
    const shown = Object.fromEntries(display.cases.filter((c) => c.numFmtId > 0).map((c) => [String(c.numFmtId), display.answers[c.id]]))
    const ids = Object.keys(run.formats)
    expect(ids.filter((id) => Number(id) <= 49).map((id) => [id, run.formats[id], run.local[id]])).toEqual(Object.entries(shown).map(([id, a]) => [id, a.code, a.local]))
    for (const id of ids.filter((x) => Number(x) >= 50)) {
      expect([id, run.formats[id], run.local[id]]).toEqual([id, run.formats[14], run.local[14]])
      expect(BUILTIN[Number(id)]).toBe(EAST_ASIAN_READ)
    }
  })
})
