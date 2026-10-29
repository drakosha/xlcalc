/**
 * Text on the clipboard as Excel pastes it (the sets `ui/paste-held`, `rows/keys` and `rows/quotes`): rows of fields at
 * tabs, quoted fields kept whole (`delimitedRows`), filled out to a rectangle as wide as the longest, so a short row's
 * missing fields are empty and clear their cells. A line break after the last row is one more row, empty, which clears
 * the cells under the paste. `allot` is given the rectangle's cells before it is made (a session's `limits.writes`).
 */
import { delimitedRows } from '../delimited.js'

export function clipboardRows(text: string, allot?: (cells: number) => void): string[][] {
  const rows = delimitedRows(text, '\t')
  const width = rows.reduce((n, r) => Math.max(n, r.length), 0)
  allot?.(rows.length * width)
  for (const r of rows) while (r.length < width) r.push('')
  return rows
}
