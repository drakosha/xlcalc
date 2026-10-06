/**
 * `restructure(book, op)`: one structure operation on a book, as a pure function. Input of the wrong
 * shape throws `InputError`; an operation Excel would refuse returns `{ ok: false, code, message }`.
 */
import type { BookInput, CellRef, RestructureResult, StructureOp } from '../contract.js'
import { checkAddress, checkBook, checkString, InputError } from '../input.js'
import { colToIndex, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { parseRange } from './area.js'
import { copyRange } from './copy.js'
import { moveRange } from './move.js'
import { shift } from './shift.js'
import { deleteSheet, renameSheet } from './sheets.js'

function object(v: unknown, path: string): Record<string, unknown> {
  if (typeof v !== 'object' || v === null || Array.isArray(v)) throw new InputError(path, `expected an object, got ${Array.isArray(v) ? 'an array' : String(v)}`)
  return v as Record<string, unknown>
}

/** The operation's fields: `op` and the ones each operation takes. */
const FIELDS: Record<string, string[]> = {
  insertRows: ['sheet', 'row', 'count'],
  deleteRows: ['sheet', 'row', 'count'],
  insertColumns: ['sheet', 'column', 'count'],
  deleteColumns: ['sheet', 'column', 'count'],
  moveRange: ['from', 'to'],
  copyRange: ['from', 'to'],
  renameSheet: ['sheet', 'name'],
  deleteSheet: ['sheet'],
}

function fields(o: Record<string, unknown>, path: string, allowed: string[]) {
  for (const k of Object.keys(o)) if (!allowed.includes(k)) throw new InputError(path, `unknown field ${JSON.stringify(k)} (expected ${allowed.join(', ')})`)
}

function integer(v: unknown, path: string, min: number, max: number): number {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < min || v > max) throw new InputError(path, `expected an integer from ${min} to ${max}, got ${JSON.stringify(v)}`)
  return v
}

/** Restructures a book; the input book is not changed. */
export function restructure(book: BookInput, op: StructureOp): RestructureResult {
  checkBook(book)
  const r = applyOp(book, op)
  if (!r.ok) return r
  // mapCell takes cells of the old book: a sheet of it (in any case) and an address
  const mapCell = (ref: CellRef): CellRef | null => {
    const x = object(ref, 'ref')
    return r.mapCell({ sheet: sheetOf(book, x.sheet, 'ref.sheet'), cell: checkAddress(x.cell, 'ref.cell') })
  }
  return { ok: true, book: r.book, mapCell }
}

/** A sheet of the book by name in any case, as the book spells it. */
function sheetOf(book: BookInput, v: unknown, path: string): string {
  const k = sheetKey(checkString(v, path))
  const sheet = book.sheets.find((s) => sheetKey(s.name) === k)
  if (!sheet) throw new InputError(path, `unknown sheet ${JSON.stringify(v)}`)
  return sheet.name
}

/** The operation on a book already checked; its mapCell takes a sheet as the book spells it and a valid address. */
export function applyOp(book: BookInput, op: StructureOp): RestructureResult {
  const o = object(op, 'op')
  const allowed = FIELDS[o.op as string]
  if (typeof o.op !== 'string' || !allowed) throw new InputError('op.op', `expected one of ${Object.keys(FIELDS).join(', ')}, got ${JSON.stringify(o.op)}`)
  fields(o, 'op', ['op', ...allowed])
  const named = (v: unknown, path: string) => sheetOf(book, v, path)
  const rangeOf = (v: unknown, path: string) => {
    const r = parseRange(checkString(v, path))
    if (!r) throw new InputError(path, `expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got ${JSON.stringify(v)}`)
    return r
  }
  const count = (max: number) => (o.count === undefined ? 1 : integer(o.count, 'op.count', 1, max))
  let r: RestructureResult
  switch (o.op as StructureOp['op']) {
    case 'insertRows':
    case 'deleteRows': {
      const sheet = named(o.sheet, 'op.sheet')
      const row = integer(o.row, 'op.row', 1, MAX_ROWS)
      const n = count(MAX_ROWS - row + 1)
      r = shift(book, sheet, 'r', row - 1, o.op === 'insertRows' ? n : -n)
      break
    }
    case 'insertColumns':
    case 'deleteColumns': {
      const sheet = named(o.sheet, 'op.sheet')
      const letters = checkString(o.column, 'op.column')
      const col = /^[A-Z]{1,3}$/.test(letters) ? colToIndex(letters) : MAX_COLS
      if (col >= MAX_COLS) throw new InputError('op.column', `expected column letters from A to XFD, got ${JSON.stringify(letters)}`)
      const n = count(MAX_COLS - col)
      r = shift(book, sheet, 'c', col, o.op === 'insertColumns' ? n : -n)
      break
    }
    case 'moveRange':
    case 'copyRange': {
      const from = object(o.from, 'op.from')
      const to = object(o.to, 'op.to')
      fields(from, 'op.from', ['sheet', 'range'])
      fields(to, 'op.to', ['sheet', o.op === 'copyRange' ? 'range' : 'cell'])
      const fromSheet = named(from.sheet, 'op.from.sheet')
      const src = rangeOf(from.range, 'op.from.range')
      const toSheet = named(to.sheet, 'op.to.sheet')
      if (o.op === 'copyRange') r = copyRange(book, fromSheet, src, toSheet, rangeOf(to.range, 'op.to.range'))
      else {
        const at = parseAddress(checkAddress(to.cell, 'op.to.cell'))
        r = moveRange(book, fromSheet, src, toSheet, at.row, at.col)
      }
      break
    }
    case 'renameSheet':
      r = renameSheet(book, named(o.sheet, 'op.sheet'), checkString(o.name, 'op.name'))
      break
    case 'deleteSheet':
      r = deleteSheet(book, named(o.sheet, 'op.sheet'))
      break
  }
  return r
}
