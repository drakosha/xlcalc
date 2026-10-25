/**
 * Excel's outline commands as structure operations (`editStructure`): Group and Ungroup of rows or columns, Show and Hide
 * Detail at a line, the level buttons; each changes one axis of one sheet as `outline-command.ts` says, or is refused as
 * Excel refuses it, with a message in the book's locale.
 */
import type { BookInput } from '../contract.js'
import { say } from '../locale/messages.js'
import type { Applied } from './geometry.js'
import { applyOutline } from './outline-command.js'
import { sheetIndex } from './rewrite.js'

export { OUTLINE_FIELDS } from './outline-command.js'

/** The operation on the sheet `sheet` of a book already checked, its other fields checked here. */
export function outlineOp(book: BookInput, sheet: string, o: Record<string, unknown>): Applied {
  const at = sheetIndex(book, sheet)
  const r = applyOutline(book.sheets[at], o)
  if (!r.ok) return { ok: false, code: r.code, message: say(book.locale, r.message) }
  if (r.sheet === book.sheets[at]) return { ok: true, book, geometry: [] }
  const sheets = book.sheets.slice()
  sheets[at] = r.sheet
  return { ok: true, book: { ...book, sheets }, geometry: [] }
}
