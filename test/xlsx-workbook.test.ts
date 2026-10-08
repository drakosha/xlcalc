/**
 * The files of the workbook sets, as the harness writes them for Excel (`bench/coverage/case-book.ts`: ExcelJS, then
 * hidden rows, an autofilter's criteria and the filter mark, `date1904` and `calcPr iterate` by hand), read into the
 * books the engine is checked with against Excel's answers (`test/workbook-cases.ts`, `test/workbook-excel.test.ts`):
 * the same hidden rows and columns, filter mode, date system and iterative calculation, and the same values.
 */
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { cellsOf, writeCaseBook } from '../bench/coverage/case-book.js'
import type { WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { compute, type BookInput, type ComputeResult } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { bookOf, CLOCK } from './workbook-cases.js'

const SETS = ['hidden', 'hidden-held', 'date1904', 'date1904-held', 'date1904-text', 'iterate', 'iterate-held']
const dir = mkdtempSync(join(tmpdir(), 'xlsx-workbook-'))
afterAll(() => rmSync(dir, { recursive: true, force: true }))

/** What a book holds beyond its cells, on the sheets of the case; Excel's defaults of iteration filled in. */
function layout(book: BookInput, sheets: string[]) {
  return {
    dateSystem: book.dateSystem ?? 1900,
    iteration: book.iteration && { maxIterations: 100, maxChange: 0.001, ...book.iteration },
    sheets: book.sheets
      .filter((s) => sheets.includes(s.name))
      .map((s) => ({
        name: s.name,
        filtered: !!s.filtered,
        rows: s.rows,
        // ExcelJS gives a hidden column its default width, 9
        columns: s.columns && Object.fromEntries(Object.entries(s.columns).map(([k, v]) => [k, v.hidden && v.width === 9 ? { ...v, width: undefined } : v])),
      })),
  }
}

const valuesAt = (r: ComputeResult, c: WorkbookCase) =>
  c.read.flatMap((g) => cellsOf(g.range).addrs.map((a) => [`${g.sheet}!${a}`, r.sheets.find((s) => s.name === g.sheet)!.cells[a]?.value ?? null]))

for (const set of SETS) {
  const cases = ((await import(`../bench/coverage/workbook/${set}.ts`)) as { buildWorkbookCases: () => WorkbookCase[] }).buildWorkbookCases()
  describe(`xlsx: the files of the workbook set ${set}`, () => {
    for (const c of cases)
      it(`${c.id}: read as the book checked against Excel`, async () => {
        const file = join(dir, `${set}-${c.id}.xlsx`)
        await writeCaseBook(c.book, c.read, file, c.file)
        const read = (await readXlsx(readFileSync(file))).book
        const book = bookOf(c)
        const sheets = book.sheets.map((s) => s.name)
        expect(JSON.parse(JSON.stringify(layout(read, sheets)))).toEqual(JSON.parse(JSON.stringify(layout(book, sheets))))
        expect(valuesAt(compute(read, { clock: CLOCK, display: false }), c)).toEqual(valuesAt(compute(book, { clock: CLOCK, display: false }), c))
      })
  })
}
