/**
 * Renaming, deleting, inserting, moving, hiding and showing a sheet. A rename takes a name by Excel's rules and rewrites every reference
 * to the sheet and the scope of its names; a delete turns references to it into `#REF!` followed by the reference
 * (`#REF!A1`, as Excel writes it), drops the names local to it and keeps the global ones with their reference lost. A
 * reference across sheets names its ends and covers the sheets between them: a new sheet changes no formula, a
 * deleted end and an end moved past the other end give way to the next sheet inward (`struct/cube`). A hidden sheet is renamed, deleted
 * and moved as any other, a very hidden one renamed alone; the book's last visible sheet is neither deleted nor hidden, as
 * Excel refuses all these (`hidden/sheets-held`).
 */
import type { BookInput } from '../contract.js'
import { sheetNameProblem } from '../input.js'
import { say, type Message } from '../locale/messages.js'
import { sheetKey } from '../syntax.js'
import type { Applied } from './geometry.js'
import { rewriteCells, rewriteNames, rewriteTemplates, sheetIndex, sheetNames, type Transform } from './rewrite.js'
import { rewriteConditional } from '../cond/structure.js'
import { rewriteValidations } from '../validation/structure.js'

export function renameSheet(book: BookInput, sheet: string, name: string): Applied {
  const key = sheetKey(sheet)
  const problem = nameProblem(book, name, sheet)
  if (problem) return { ok: false, code: 'sheet_name', message: say(book.locale, problem) }
  const renamed = (s: string) => (sheetKey(s) === key ? name : s)
  const t: Transform = {
    ref: ({ sheet: s, area }) => (sheetKey(s) === key ? { sheet: name, area } : undefined),
    sheet: (s) => (sheetKey(s) === key ? name : undefined),
    ends: (a, b) => (sheetKey(a) === key || sheetKey(b) === key ? [renamed(a), renamed(b)] : undefined),
  }
  const out = rewriteBook(book, t, sheetNames(book), (s) => (sheetKey(s) === key ? name : s))
  if (out.activeSheet !== undefined) out.activeSheet = renamed(out.activeSheet)
  if (out.names) out.names = out.names.map((n) => (n.scope !== undefined && sheetKey(n.scope) === key ? { ...n, scope: name } : n))
  return { ok: true, book: out, geometry: [{ kind: 'renameSheet', s: sheetIndex(book, sheet) }] }
}

/**
 * The ends of references across sheets when the sheet `key` leaves their span (deleted, or an end moved away): an
 * end on it goes to the next sheet inward in the book's order; nothing left is null.
 */
function endsWithout(book: BookInput, key: string): Transform['ends'] {
  const order = book.sheets.map((s) => s.name)
  const at = (s: string) => order.findIndex((n) => sheetKey(n) === sheetKey(s))
  return (a, b) => {
    // an end that is no sheet of the book leaves the reference as written: it is #REF! already
    if (at(a) < 0 || at(b) < 0) return undefined
    let [lo, hi] = [Math.min(at(a), at(b)), Math.max(at(a), at(b))]
    if (sheetKey(order[lo]) !== key && sheetKey(order[hi]) !== key) return undefined
    if (sheetKey(order[lo]) === key) lo++
    else hi--
    return lo > hi ? null : [order[lo], order[hi]]
  }
}

/** Excel's rules for a new sheet name, which no sheet but `own` has; null when it is fine. */
function nameProblem(book: BookInput, name: string, own?: string): Message | null {
  const taken = book.sheets.some((s) => sheetKey(s.name) === sheetKey(name) && (own === undefined || sheetKey(s.name) !== sheetKey(own)))
  return sheetNameProblem(name) ?? (taken ? { key: 'sheetNameTaken', name } : null)
}

/** A new empty sheet before `before`, else at the end; references across sheets take it in between their ends. */
export function insertSheet(book: BookInput, name: string, before: string | undefined): Applied {
  const problem = nameProblem(book, name)
  if (problem) return { ok: false, code: 'sheet_name', message: say(book.locale, problem) }
  const at = before === undefined ? book.sheets.length : sheetIndex(book, before)
  const sheets = [...book.sheets]
  sheets.splice(at, 0, { name, cells: {} })
  return { ok: true, book: { ...book, sheets }, geometry: [{ kind: 'insertSheet', at }] }
}

/**
 * A sheet moved before `before`, else to the end; references across sheets cover what lies between their ends then.
 * Excel places it by the visible sheets (`hidden/sheets-held`, `sheets-held-2`): before a hidden sheet it goes before the
 * next visible one, or after the last visible one when none follows, and the end is after the last visible sheet. An
 * end of a reference across sheets that moves stays its end while it does not pass the other end; one that passes it
 * gives way to the next sheet inward.
 */
export function moveSheet(book: BookInput, sheet: string, before: string | undefined): Applied {
  const key = sheetKey(sheet)
  const s = sheetIndex(book, sheet)
  const moving = book.sheets[s]
  if (moving.hidden === 'very') return veryHidden(book, moving.name)
  if (before !== undefined && sheetKey(before) === key) return { ok: true, book, geometry: [{ kind: 'moveSheet', s, at: s }] }
  const sheets = book.sheets.filter((x) => x !== moving)
  const at = movedTo(sheets, before)
  sheets.splice(at, 0, moving)
  const [was, now] = [book.sheets, sheets].map((list) => list.map((s) => sheetKey(s.name)))
  const side = (order: string[], a: string, b: string) => Math.sign(order.indexOf(sheetKey(a)) - order.indexOf(sheetKey(b)))
  const without = endsWithout(book, key)!
  const t: Transform = { ref: () => undefined, ends: (a, b) => (sheetKey(a) === sheetKey(b) || side(was, a, b) === side(now, a, b) ? undefined : without(a, b)) }
  return { ok: true, book: rewriteBook({ ...book, sheets }, t, sheetNames(book)), geometry: [{ kind: 'moveSheet', s, at }] }
}

/** Where a sheet moved before `before` (else to the end) goes among the other sheets: by the visible ones, as Excel places it. */
function movedTo(sheets: BookInput['sheets'], before: string | undefined): number {
  const at = before === undefined ? sheets.length : sheetIndex({ sheets }, before)
  const next = sheets.findIndex((s, i) => i >= at && !s.hidden)
  if (next >= 0) return next
  const last = sheets.findLastIndex((s) => !s.hidden)
  return last >= 0 ? last + 1 : at
}

export function deleteSheet(book: BookInput, sheet: string): Applied {
  if (book.sheets.length === 1) return { ok: false, code: 'last_sheet', message: say(book.locale, { key: 'lastSheet' }) }
  const key = sheetKey(sheet)
  const s = sheetIndex(book, sheet)
  if (book.sheets[s].hidden === 'very') return veryHidden(book, sheet)
  if (lastVisible(book, key)) return { ok: false, code: 'last_sheet', message: say(book.locale, { key: 'lastVisibleSheet', sheet }) }
  const t: Transform = { ref: ({ sheet: s }) => (sheetKey(s) === key ? 'gone' : undefined), sheet: (s) => (sheetKey(s) === key ? null : undefined), ends: endsWithout(book, key) }
  const rest: BookInput = { ...book, sheets: book.sheets.filter((_, i) => i !== s) }
  if (rest.activeSheet !== undefined && sheetKey(rest.activeSheet) === key) rest.activeSheet = rest.sheets.find((s) => !s.hidden)!.name
  if (book.names) rest.names = book.names.filter((n) => n.scope === undefined || sheetKey(n.scope) !== key)
  return { ok: true, book: rewriteBook(rest, t, sheetNames(book)), geometry: [{ kind: 'deleteSheet', s }] }
}

/** Excel deletes and moves no very hidden sheet: its object model answers an error, Excel for Mac at times nothing. */
const veryHidden = (book: BookInput, sheet: string): Applied => ({ ok: false, code: 'very_hidden', message: say(book.locale, { key: 'veryHidden', sheet }) })

/** Whether the sheet `key` is the book's only visible sheet. */
const lastVisible = (book: BookInput, key: string) => book.sheets.every((s) => (sheetKey(s.name) === key) === !s.hidden)

/**
 * The sheet hidden (very hidden with `very`) or shown again (`hidden` false); no formula changes, as formulas read a
 * hidden sheet as any other. The book's last visible sheet is not hidden, as Excel refuses.
 */
export function hideSheet(book: BookInput, sheet: string, hidden: boolean | 'very'): Applied {
  const key = sheetKey(sheet)
  if (hidden && lastVisible(book, key)) return { ok: false, code: 'last_sheet', message: say(book.locale, { key: 'lastVisibleSheet', sheet }) }
  const sheets: BookInput['sheets'] = book.sheets.map((s) => {
    if (sheetKey(s.name) !== key) return s
    const { hidden: _, ...rest } = s
    return hidden ? { ...rest, hidden } : rest
  })
  const activeSheet = hidden && book.activeSheet !== undefined && sheetKey(book.activeSheet) === key ? sheets.find((s) => !s.hidden)!.name : book.activeSheet
  return { ok: true, book: { ...book, sheets, ...(activeSheet !== undefined ? { activeSheet } : {}) }, geometry: [] }
}

/**
 * Every formula of the book (the names and the calculated columns too) rewritten in place, for operations that move
 * no cells; `rename` gives a sheet's new name.
 */
function rewriteBook(book: BookInput, t: Transform, sheets: Map<string, string>, rename = (name: string) => name): BookInput {
  const out: BookInput = {
    ...book,
    sheets: book.sheets.map((s) => {
      const name = rename(s.name)
      const cells = rewriteCells(s.cells, s.name, name, t, sheets)
      const tables = rewriteTemplates(s.tables, s.name, () => name, t, sheets)
      const cf = rewriteConditional(s.conditionalFormats, s.name, name, t, sheets)
      const validations = rewriteValidations(s.validations, s.name, name, t, sheets)
      return cells !== s.cells || name !== s.name || tables !== s.tables || cf !== s.conditionalFormats || validations !== s.validations
        ? { ...s, name, cells, ...(tables ? { tables } : {}), ...(cf ? { conditionalFormats: cf } : {}), ...(validations ? { validations } : {}) }
        : s
    }),
  }
  const names = rewriteNames(book.names, t, sheets)
  return names ? { ...out, names } : out
}
