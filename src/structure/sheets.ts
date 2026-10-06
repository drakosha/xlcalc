/**
 * Renaming and deleting a sheet. A rename takes a name by Excel's rules and rewrites every reference to the
 * sheet and the scope of its names; a delete turns references to it into `#REF!` followed by the reference (`#REF!A1`, as Excel
 * writes it), drops the names local to it and keeps the global ones with their reference lost.
 */
import type { BookInput, CellRef, RestructureResult } from '../contract.js'
import { sheetNameProblem } from '../input.js'
import { sheetKey } from '../syntax.js'
import { rewriteBook, sheetNames, type Transform } from './rewrite.js'

export function renameSheet(book: BookInput, sheet: string, name: string): RestructureResult {
  const key = sheetKey(sheet)
  const problem = sheetNameProblem(name) ?? (book.sheets.some((s) => sheetKey(s.name) === sheetKey(name) && sheetKey(s.name) !== key) ? `there is a sheet named ${JSON.stringify(name)} already (sheet names are case-insensitive)` : null)
  if (problem) return { ok: false, code: 'sheet_name', message: problem }
  const t: Transform = {
    ref: ({ sheet: s, area }) => (sheetKey(s) === key ? { sheet: name, area } : undefined),
    sheet: (s) => (sheetKey(s) === key ? name : undefined),
  }
  const out = rewriteBook(book, t, sheetNames(book), (s) => (sheetKey(s) === key ? name : s))
  if (out.names) out.names = out.names.map((n) => (n.scope !== undefined && sheetKey(n.scope) === key ? { ...n, scope: name } : n))
  const mapCell = (ref: CellRef): CellRef => (sheetKey(ref.sheet) === key ? { sheet: name, cell: ref.cell } : ref)
  return { ok: true, book: out, mapCell }
}

export function deleteSheet(book: BookInput, sheet: string): RestructureResult {
  if (book.sheets.length === 1) return { ok: false, code: 'last_sheet', message: 'a book keeps at least one sheet' }
  const key = sheetKey(sheet)
  const t: Transform = { ref: ({ sheet: s }) => (sheetKey(s) === key ? 'gone' : undefined), sheet: (s) => (sheetKey(s) === key ? null : undefined) }
  const rest: BookInput = { ...book, sheets: book.sheets.filter((s) => sheetKey(s.name) !== key) }
  if (book.names) rest.names = book.names.filter((n) => n.scope === undefined || sheetKey(n.scope) !== key)
  const mapCell = (ref: CellRef): CellRef | null => (sheetKey(ref.sheet) === key ? null : ref)
  return { ok: true, book: rewriteBook(rest, t, sheetNames(book)), mapCell }
}
