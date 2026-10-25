/**
 * `editStructure(book, op)`: one structure operation on a book, as a pure function. Input of the wrong
 * shape throws `InputError`; an operation Excel would refuse returns `{ ok: false, code, message }`.
 */
import type { BookInput, CellRef, FillType, StructureResult, StructureOp } from '../contract.js'
import { checkBook, checkSheetOf } from '../input.js'
import { InputError } from '../input-error.js'
import { checkOneOf, checkString, integer, isObject, onlyKeys } from '../input/basic.js'
import { checkAddress, checkRange } from '../input/range.js'
import { opTaken } from '../taken.js'
import { colToIndex, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { say } from '../locale/messages.js'
import { createState } from '../engine.js'
import { currentClock } from '../env.js'
import { clockSerial } from '../result.js'
import { areaText } from './area.js'
import { looksTaken } from '../looks/sheet.js'
import { linesMoved, overLooks } from './looks.js'
import { GEOMETRY, mapCellOf, type Applied, type SentGeometry } from './geometry.js'
import { withGrids } from './grid.js'
import { arrayInTable } from './arrays.js'
import { copyRange } from './copy.js'
import { copyRows, deleteRows, insertRows } from './filtered.js'
import { moveRange } from './move.js'
import { shift } from './shift.js'
import { sheetIndex } from './rewrite.js'
import { deleteSheet, hideSheet, insertSheet, moveSheet, renameSheet } from './sheets.js'
import { stateValues, type ValuesIn } from './values.js'
import { checkWrites } from './writes.js'
import { tableOp } from '../tables/ops.js'
import { OUTLINE_FIELDS, outlineOp } from './outline.js'
import { mergeCells, unmergeCells } from './merge.js'
import { fillRange, type Copier } from './fill.js'
import { copyTables, moveTables } from '../tables/moves.js'
import { rewriteBookTables } from '../tables/rewrite.js'

function object(v: unknown, path: string): Record<string, unknown> {
  if (!isObject(v)) throw new InputError(path, `expected an object, got ${Array.isArray(v) ? 'an array' : String(v)}`)
  return v
}

/** The operation's fields: `op` and the ones each operation takes. */
const FIELDS: Record<string, string[]> = {
  insertRows: ['sheet', 'row', 'count'],
  deleteRows: ['sheet', 'row', 'count'],
  insertColumns: ['sheet', 'column', 'count'],
  deleteColumns: ['sheet', 'column', 'count'],
  moveRange: ['from', 'to'],
  copyRange: ['from', 'to', 'withMeta'],
  renameSheet: ['sheet', 'name'],
  deleteSheet: ['sheet'],
  insertSheet: ['name', 'before'],
  moveSheet: ['sheet', 'before'],
  hideSheet: ['sheet', 'very'],
  unhideSheet: ['sheet'],
  addTable: ['sheet', 'table'],
  renameTable: ['table', 'name'],
  deleteTable: ['table'],
  convertTable: ['table'],
  ...OUTLINE_FIELDS,
  mergeCells: ['sheet', 'range', 'across', 'center'],
  unmergeCells: ['sheet', 'range'],
  fillRange: ['sheet', 'range', 'to', 'type'],
}

const FILL_TYPES: readonly FillType[] = ['default', 'series', 'copy']

/**
 * Restructures a workbook; the input workbook is not modified. A copy that pastes values (from a filtered sheet)
 * calculates the workbook at the current time.
 */
export function editStructure(book: BookInput, op: StructureOp): StructureResult {
  book = checkBook(book)
  op = opTaken(op)
  const r = applyOp(book, op, lazyValues(book))
  if (!r.ok) return r
  const out = { ...r.book, sheets: r.book.sheets.map((s) => looksTaken(s)) }
  return { ok: true, book: out, mapCell: checkedMapCell({ geometry: r.geometry, before: namesOf(book), after: namesOf(out) }) }
}

/** The names of a book's sheets, in order. */
export const namesOf = (book: { sheets: { name: string }[] }) => book.sheets.map((s) => s.name)

/**
 * The mapCell callers get, made from a geometry: it takes cells of the book before (a sheet's name, in any case, and an
 * address) and checks their shape; a sheet the book before does not have holds no cell of it, so its cells are null, as
 * a host's window can hold a sheet that is gone. It is kept with its geometry, which a pool's thread sends instead of
 * the function (`GEOMETRY`).
 */
export function checkedMapCell(g: SentGeometry): (ref: CellRef) => CellRef | null {
  const inner = mapCellOf(g.geometry, g.before, g.after)
  const mapCell = (ref: CellRef) => {
    const x = object(ref, 'ref')
    return inner({ sheet: checkString(x.sheet, 'ref.sheet'), cell: checkAddress(x.cell, 'ref.cell') })
  }
  GEOMETRY.set(mapCell, g)
  return mapCell
}

/** The values of a book computed when first asked for, at the current moment. */
function lazyValues(book: BookInput): ValuesIn {
  let values: ValuesIn | undefined
  return (sheet, rect) => {
    if (!values) {
      const clock = currentClock()
      values = stateValues(createState(book, { now: clockSerial(clock), seed: clock.seed }))
    }
    return values(sheet, rect)
  }
}

/**
 * The operation on a book already checked, with its geometry. `values` gives the computed values a paste of values takes.
 * Its steps read each sheet's cells once (`withGrids`); a copy or a move over `limits.writes` stops first. The sheets'
 * looks move with their lines as the lines do; those of the cells it copies, moves or writes otherwise are given to it as
 * entries (`overLooks`), so the book it gives may hold entries of a look alone: the engine takes it as it takes a host's.
 * `whole` is a copy of the fill handle, which takes its source whole and lays it over the rows a filter hid too.
 */
export const applyOp = (book: BookInput, op: StructureOp, values: ValuesIn, whole = false): Applied =>
  withGrids(() => {
    checkWrites(book, op, values, whole)
    const input = overLooks(book, op)
    const r = applied(input, op, values, whole)
    return r.ok ? linesMoved(input, r) : r
  })

function applied(book: BookInput, op: StructureOp, values: ValuesIn, whole: boolean): Applied {
  const o = object(op, 'op')
  // an own field only: a name of Object.prototype's members is no operation
  const allowed = typeof o.op === 'string' && Object.hasOwn(FIELDS, o.op) ? FIELDS[o.op] : undefined
  if (typeof o.op !== 'string' || !allowed) throw new InputError('op.op', `expected one of ${Object.keys(FIELDS).join(', ')}, got ${JSON.stringify(o.op)}`)
  onlyKeys(o, 'op', ['op', ...allowed])
  const named = (v: unknown, path: string) => checkSheetOf(book, v, path).name
  const count = (max: number) => (o.count === undefined ? 1 : integer(o.count, 'op.count', 1, max))
  let r: Applied
  switch (o.op as StructureOp['op']) {
    case 'insertRows':
    case 'deleteRows': {
      const sheet = named(o.sheet, 'op.sheet')
      const row = integer(o.row, 'op.row', 1, MAX_ROWS)
      const n = count(MAX_ROWS - row + 1)
      r = (o.op === 'insertRows' ? insertRows : deleteRows)(book, sheet, row - 1, n)
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
      onlyKeys(from, 'op.from', ['sheet', 'range'])
      onlyKeys(to, 'op.to', ['sheet', o.op === 'copyRange' ? 'range' : 'cell'])
      const fromSheet = named(from.sheet, 'op.from.sheet')
      const src = checkRange(from.range, 'op.from.range')
      const toSheet = named(to.sheet, 'op.to.sheet')
      if (o.withMeta !== undefined && typeof o.withMeta !== 'boolean') throw new InputError('op.withMeta', `expected a boolean, got ${JSON.stringify(o.withMeta)}`)
      // the cells as pasted, before a header cell pasted over takes its value as its column's name
      let pasted: Applied
      if (o.op === 'copyRange') {
        const target = checkRange(to.range, 'op.to.range')
        const withMeta = o.withMeta === true
        pasted = whole ? copyRange(book, fromSheet, src, toSheet, target, withMeta) : copyRows(book, fromSheet, src, toSheet, target, withMeta, values)
        r = copyTables(book, fromSheet, src, toSheet, target).after(pasted)
      } else {
        const at = parseAddress(checkAddress(to.cell, 'op.to.cell'))
        const tables = moveTables(book, fromSheet, src, toSheet, at.row - src.r1, at.col - src.c1)
        if (tables.refuse) pasted = r = { ok: false, code: 'table', message: say(book.locale, tables.refuse) }
        else {
          pasted = moveRange(tables.before(book), fromSheet, src, toSheet, at.row, at.col)
          r = pasted.ok ? { ...pasted, book: tables.after(pasted.book) } : pasted
        }
      }
      if (r.ok && pasted.ok) {
        // an array of several cells over a table, its header too, against the tables where the operation leaves them
        const ti = sheetIndex(book, toSheet)
        const refused = arrayInTable({ ...pasted.book.sheets[ti], tables: r.book.sheets[ti].tables })
        if (refused) r = { ok: false, code: 'table', message: say(book.locale, refused) }
      }
      break
    }
    case 'renameSheet':
      r = renameSheet(book, named(o.sheet, 'op.sheet'), checkString(o.name, 'op.name'))
      break
    case 'deleteSheet': {
      // references to the sheet's tables are lost with it
      const sheet = named(o.sheet, 'op.sheet')
      const gone = sheetIndex(book, sheet)
      r = deleteSheet(book.sheets.length > 1 ? rewriteBookTables(book, (t) => (t.s === gone ? 'lost' : undefined)) : book, sheet)
      break
    }
    case 'insertSheet':
      r = insertSheet(book, checkString(o.name, 'op.name'), o.before === undefined ? undefined : named(o.before, 'op.before'))
      break
    case 'moveSheet':
      r = moveSheet(book, named(o.sheet, 'op.sheet'), o.before === undefined ? undefined : named(o.before, 'op.before'))
      break
    case 'hideSheet':
    case 'unhideSheet': {
      const sheet = named(o.sheet, 'op.sheet')
      if (o.very !== undefined && typeof o.very !== 'boolean') throw new InputError('op.very', `expected a boolean, got ${JSON.stringify(o.very)}`)
      r = hideSheet(book, sheet, o.op === 'unhideSheet' ? false : o.very ? 'very' : true)
      break
    }
    case 'addTable':
    case 'renameTable':
    case 'deleteTable':
    case 'convertTable':
      r = tableOp(book, op as Extract<StructureOp, { op: 'addTable' | 'renameTable' | 'deleteTable' | 'convertTable' }>)
      break
    case 'group':
    case 'ungroup':
    case 'showDetail':
    case 'hideDetail':
    case 'showLevels':
      r = outlineOp(book, named(o.sheet, 'op.sheet'), o)
      break
    case 'mergeCells':
    case 'unmergeCells': {
      const sheet = named(o.sheet, 'op.sheet')
      const range = checkRange(o.range, 'op.range')
      const flag = (k: 'across' | 'center') => {
        if (o[k] !== undefined && typeof o[k] !== 'boolean') throw new InputError(`op.${k}`, `expected a boolean, got ${JSON.stringify(o[k])}`)
        return o[k] === true
      }
      r = o.op === 'mergeCells' ? mergeCells(book, sheet, range, flag('across'), flag('center')) : unmergeCells(book, sheet, range)
      break
    }
    case 'fillRange': {
      if (o.type !== undefined) checkOneOf(o.type, 'op.type', FILL_TYPES)
      const sheet = named(o.sheet, 'op.sheet')
      // the fill handle is no copy and paste: Excel's fills the rows a filter hid too (the struct set `hidden`)
      const copy: Copier = (b, from, at) => applyOp(b, { op: 'copyRange', from: { sheet, range: areaText(from) }, to: { sheet, range: areaText(at) } }, values, true)
      r = fillRange(book, sheet, checkRange(o.range, 'op.range'), checkRange(o.to, 'op.to'), String(o.to), (o.type ?? 'default') as FillType, copy)
      break
    }
  }
  return r
}
