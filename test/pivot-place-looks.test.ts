/**
 * The look Excel keeps in the cells of a PivotTable's area, written as `placePivots` lays them out: a compact form's row
 * labels left aligned and indented by their level, the other cells without an alignment of their own; against Excel's
 * books (`test/excel/win/pivot/<set>/`), each PivotTable laid out again from what Excel saved.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { formatAddress } from '../src/model.js'
import { layoutPivot } from '../src/pivot/layout.js'
import { placePivots } from '../src/pivot/place.js'
import { readXlsx } from '../src/xlsx.js'
import { savedResult } from './pivot-saved.js'

const ROOT = new URL('./excel/win/pivot/', import.meta.url).pathname
const SETS = ['basic', 'layout', 'filter', 'sort', 'edit', 'layout-held', 'filter-held'].filter((s) => existsSync(`${ROOT}${s}.json`))

describe.each(SETS)('the alignment of the cells of the PivotTables of set %s', (set) => {
  const books = readdirSync(`${ROOT}${set}`).filter((f) => f.endsWith('.xlsx'))
  it.each(books)('%s', async (file) => {
    const { book } = await readXlsx(readFileSync(`${ROOT}${set}/${file}`))
    for (const sheet of book.sheets)
      for (const t of sheet.pivotTables ?? []) {
        const cache = book.pivotCaches!.find((c) => c.id === t.cache)!
        const laid = layoutPivot(savedResult(t, cache, sheet), book)
        // the cells as Excel saved them, without their own looks: the layout gives them
        const bare = { ...book, sheets: book.sheets.map((s) => (s === sheet ? { ...s, cells: Object.fromEntries(Object.entries(s.cells).map(([a, c]) => [a, { ...c, style: undefined }])) } : s)) }
        const r = placePivots(bare, [{ sheet: sheet.name, table: t, laid }], 'spill')
        if (!r.ok) throw new Error(r.message)
        const of = (b: typeof book, s: number, a: string) => {
          const st = b.sheets[s].cells[a]?.style
          return st === undefined ? undefined : b.styles?.[st]?.alignment
        }
        const s = book.sheets.indexOf(sheet)
        const got = laid.cells.map((c) => [formatAddress(c.row, c.col), of(r.book, s, formatAddress(c.row, c.col))])
        const want = laid.cells.map((c) => [formatAddress(c.row, c.col), of(book, s, formatAddress(c.row, c.col))])
        expect(got, `${sheet.name}!${t.name}`).toEqual(want)
      }
  })
})
