/**
 * References in a file's attributes as Excel reads them: the ranges of a rule (`sqref` of conditional formats and data
 * validation). `$` and the letters' case do not count (`$a$1:B2` is A1:B2): Excel for Mac and for Windows take each so
 * and save it upper case without `$` (the sets `cond/sqref-held`, `valid/sqref-held`).
 */
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
