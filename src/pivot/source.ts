/**
 * A pivot cache's source in the workbook as cells: a range of a sheet, a table (its headers and data rows), or a defined
 * name over either, a dynamic one (`OFFSET(…)`) by the reference its formula gives.
 */
import type { BookInput, PivotSource } from '../contract.js'
import { sheetLines } from '../layout/lines.js'
import { shownAs } from '../numfmt/shown-as.js'
import { sourcePlace, type SourceRefusal } from '../pivot-source.js'
import type { ValuesIn } from '../structure/values.js'
import type { PivotSourceGrid } from './refresh.js'

export type { SourceRefusal }

/**
 * A source's cells as a grid, by the computed values `values`: the header row first, a range that runs past the sheet's
 * last cell in it cut there; a cell shows a date by its number format, and each column's format is its first row's.
 */
export function sourceGrid(book: BookInput, source: PivotSource, values: ValuesIn): PivotSourceGrid | SourceRefusal {
  const place = sourcePlace(book, source)
  if (typeof place === 'string') return place
  const { s, rect } = place
  const sheet = book.sheets[s]
  const got = values(sheet.name, rect)
  const last = got.reduce((m, [r]) => Math.max(m, r), rect.r1 + 1)
  const r2 = Math.min(rect.r2, last)
  const width = rect.c2 - rect.c1 + 1
  const rows = Array.from({ length: r2 - rect.r1 + 1 }, () => new Array(width).fill(null))
  for (const [r, c, v] of got) if (r <= r2) rows[r - rect.r1][c - rect.c1] = v
  if (place.headers) rows[0] = place.headers.slice()
  const lines = sheetLines(sheet)
  const formats = Array.from({ length: width }, (_, c) => lines.format(rect.r1 + 1, rect.c1 + c))
  return {
    rows,
    dated: (r, c) => typeof rows[r][c] === 'number' && shownAs(lines.format(rect.r1 + r, rect.c1 + c)) !== 'number',
    formats,
  }
}
