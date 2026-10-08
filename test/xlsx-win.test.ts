/**
 * The xlsx filter files as Excel for Windows computed and saved them (`test/excel/win/xlsx/filters/`, `bench/windows`):
 * read back, each has the filter mode Excel for Mac took, and the engine computes the values Excel for Windows stored,
 * as `xlsx-excel.test.ts` checks the Mac's.
 */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { XLSX_CLOCK } from '../bench/coverage/xlsx/books.js'
import { compareCached } from '../bench/coverage/xlsx/compare.js'
import { FILTER_FILES, filterFile } from '../bench/coverage/xlsx/filters.js'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { compute } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'

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
      expect(run.books[c.id].sent, 'the file changed since Excel saw it: run the xlsx filters step').toBe(createHash('sha256').update(data).digest('hex'))
      expect(run.books[c.id].rejected).toBeUndefined()
      const read = await readXlsx(readFileSync(new URL(`filters/${c.id}.xlsx`, dir)))
      const onMac = await readXlsx(readFileSync(new URL(`filters/${c.id}.xlsx`, mac)))
      expect(!!read.book.sheets[0].filtered).toBe(!!onMac.book.sheets[0].filtered)
      expect(compareCached(compute(read.book, { clock: XLSX_CLOCK }), read.cached)).toEqual([])
    })
})
