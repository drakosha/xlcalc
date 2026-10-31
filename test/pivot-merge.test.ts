/**
 * Merge and center cells with labels (`mergeLabels`) against Excel's books (set `merge-1`): each PivotTable laid out
 * again from what Excel saved gives Excel's merged areas and the alignment of each cell of its area.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { formatAddress } from '../src/model.js'
import { layoutPivot } from '../src/pivot/layout.js'
import { placePivots } from '../src/pivot/place.js'
import { readXlsx } from '../src/xlsx.js'
import { savedResult } from './pivot-saved.js'

const DIR = new URL('./excel/win/pivot/merge-1/', import.meta.url).pathname

describe('mergeLabels against Excel for Windows: the set merge-1', () => {
  it.each(readdirSync(DIR).filter((f) => f.endsWith('.xlsx')))('%s', async (file) => {
    const { book } = await readXlsx(readFileSync(DIR + file))
    const s = book.sheets.findIndex((x) => x.pivotTables)
    const sheet = book.sheets[s]
    const t = sheet.pivotTables![0]
    expect(t.mergeLabels).toBe(true)
    const laid = layoutPivot(savedResult(t, book.pivotCaches!.find((c) => c.id === t.cache)!, sheet), book)
    // the cells and merged areas as Excel saved them, without their looks: the layout gives them
    const bare = { ...book, sheets: book.sheets.map((x, i) => (i === s ? { ...x, merges: undefined, cells: Object.fromEntries(Object.entries(x.cells).map(([a, c]) => [a, { ...c, style: undefined }])) } : x)) }
    const r = placePivots(bare, [{ sheet: sheet.name, table: t, laid }], 'spill')
    if (!r.ok) throw new Error(r.message)
    const align = (b: typeof book, a: string) => {
      const st = b.sheets[s].cells[a]?.style
      return st === undefined ? undefined : b.styles?.[st]?.alignment
    }
    expect([...(r.book.sheets[s].merges ?? [])].sort()).toEqual([...(sheet.merges ?? [])].sort())
    const cells = laid.cells.map((c) => formatAddress(c.row, c.col))
    expect(cells.map((a) => [a, align(r.book, a)])).toEqual(cells.map((a) => [a, align(book, a)]))
  })
})
