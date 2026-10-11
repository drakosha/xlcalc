/**
 * References in a file's attributes as Excel reads them: the ranges of a rule (`sqref` of conditional formats and data
 * validation) and the cell of a note. `$` and the letters' case do not count and a row may have leading zeros
 * (`$a$01:B2` is A1:B2): Excel for Mac and for Windows take each so and save it upper case without `$` (the sets
 * `cond/sqref-held`, `valid/sqref-held`, `parts/refs-held`).
 */
import { formatAddress } from '../../model.js'
import { parseRange, type Rect } from '../../structure/area.js'

/** A range (`B2`, `B2:D9`, `B:D`, `2:9`) as Excel reads one in a file, or null when it is none of the sheet. */
export function fileRange(text: string): Rect | null {
  // letters of other scripts stay out: some come to Latin ones in upper case (`ß` is `SS`)
  return /^[A-Za-z0-9$:]+$/.test(text) ? parseRange(text.replaceAll('$', '').toUpperCase()) : null
}

/** The ranges of a `sqref`, separated by spaces: each as written, and as read (null for one that is no range of the sheet). */
export const sqrefRanges = (sqref: string): { text: string; rect: Rect | null }[] =>
  sqref
    .split(/\s+/)
    .filter(Boolean)
    .map((text) => ({ text, rect: fileRange(text) }))

/** The cell a reference names, as the model keys it (`A01`, `$A$1`, `a1` and `A1:A1` are A1), or null when it names no one cell. */
export function fileCell(text: string): string | null {
  const r = fileRange(text)
  return r && r.r1 === r.r2 && r.c1 === r.c2 ? formatAddress(r.r1, r.c1) : null
}
