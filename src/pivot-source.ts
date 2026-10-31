/**
 * A pivot cache's source in the workbook as a place: a range of a sheet, a table (its headers and data rows), or a defined
 * name over either, a dynamic one (`OFFSET(…)`) by the reference its formula gives. In the root, as a new PivotTable
 * takes its fields from it (`addPivotTable`); `xlcalc/pivot` reads its cells (`pivot/source.ts`).
 */
import { caseKey } from './collate.js'
import type { BookInput, PivotSource } from './contract.js'
import { createState } from './engine.js'
import { clockOptions } from './result.js'
import { parseRange, type Rect } from './structure/area.js'
import { stateValues } from './structure/values.js'
import { rowsOf, tablesOf } from './tables/model.js'

/** Why a source cannot be read: another workbook, a kind not refreshed here, or a sheet, table, name or range not there. */
export type SourceRefusal = 'other_book' | 'not_refreshed' | 'not_found'

/** A source's place: its sheet's index and rectangle, and the headers a table gives in place of a row of cells. */
export interface SourcePlace {
  s: number
  rect: Rect
  headers?: string[]
}

const sheetOf = (book: BookInput, name: string) => book.sheets.findIndex((x) => caseKey(x.name) === caseKey(name))

/** A range's text without its `$`, as a rectangle. */
const rangeOf = (text: string) => parseRange(text.replace(/\$/g, '').toUpperCase())

/** A sheet-qualified reference (`'My Data'!$A$1:$J$49`), or null. */
const REF = /^(?:'((?:[^']|'')+)'|([^!'\s()]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})$/

function tablePlace(book: BookInput, name: string): SourcePlace | null {
  const t = tablesOf(book).byName.get(caseKey(name))
  if (!t) return null
  const { first, last } = rowsOf(t)
  // the table's column names head the fields, its data rows are the records
  return { s: t.s, rect: { r1: first - 1, r2: last, c1: t.c1, c2: t.c2 }, headers: t.cols }
}

/** The reference a name's formula gives, by calculating it in the workbook: its sheet, rows and columns. */
function calculated(book: BookInput, name: string): SourcePlace | null {
  const probe = `__pivot_source_${book.sheets.length}`
  // ROW and COLUMN of a range spill: their least is its first row and column
  const cells = Object.fromEntries(['SHEET(#)', 'MIN(ROW(#))', 'MIN(COLUMN(#))', 'ROWS(#)', 'COLUMNS(#)'].map((f, i) => [`A${i + 1}`, { formula: f.replace('#', name) }]))
  const state = createState({ ...book, sheets: [...book.sheets, { name: probe, cells }] }, clockOptions({ now: '2000-01-01', seed: 0 }))
  const got = new Map(stateValues(state)(probe, { r1: 0, r2: 4, c1: 0, c2: 0 }).map(([r, , v]) => [r, v]))
  const [s, r, c, h, w] = [0, 1, 2, 3, 4].map((i) => got.get(i))
  if (![s, r, c, h, w].every((x) => typeof x === 'number')) return null
  return { s: (s as number) - 1, rect: { r1: (r as number) - 1, c1: (c as number) - 1, r2: (r as number) + (h as number) - 2, c2: (c as number) + (w as number) - 2 } }
}

function namePlace(book: BookInput, name: string): SourcePlace | null {
  const local = /^(.+)!([^!]+)$/.exec(name)
  const found = (book.names ?? []).find((n) => caseKey(n.name) === caseKey(local ? local[2] : name) && (local ? n.scope !== undefined && caseKey(n.scope) === caseKey(local[1].replace(/^'|'$/g, '')) : n.scope === undefined))
  // a file names a table's source as it names a defined name's
  if (!found) return local ? null : tablePlace(book, name)
  const ref = found.ref.replace(/^=/, '').trim()
  const m = REF.exec(ref)
  if (m) {
    const s = sheetOf(book, m[1] ? m[1].replace(/''/g, "'") : m[2])
    const rect = rangeOf(m[3])
    return s >= 0 && rect ? { s, rect } : null
  }
  return tablePlace(book, ref.replace(/\[#(All|Data)\]$/i, '')) ?? calculated(book, name)
}

/** Where a source's cells are, or why they cannot be read. */
export function sourcePlace(book: BookInput, source: PivotSource): SourcePlace | SourceRefusal {
  if ('book' in source && source.book !== undefined) return 'other_book'
  switch (source.type) {
    case 'range': {
      const s = sheetOf(book, source.sheet)
      const rect = rangeOf(source.range)
      return s >= 0 && rect ? { s, rect } : 'not_found'
    }
    case 'table':
      return tablePlace(book, source.table) ?? 'not_found'
    case 'name':
      return namePlace(book, source.name) ?? 'not_found'
  }
  return 'not_refreshed'
}

