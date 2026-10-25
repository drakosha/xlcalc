/**
 * Inserting and deleting rows and columns, with references adjusted as Excel adjusts them (checked in
 * `test/excel/struct/rows.json` and `edge.json`): a range grows when rows are inserted inside it and
 * shrinks when some of its rows are deleted, a deleted reference becomes `#REF!`, absolute references
 * move too. At the sheet's edge a range ending on the last row keeps ending there, a single cell pushed
 * past it is `#REF!`, and a range over every row is not changed by rows inserted or deleted.
 */
import type { BookInput, CellInput, SheetInput } from '../contract.js'
import { formatAddress } from '../model.js'
import { say, type EngineMessage } from '../locale/messages.js'
import { sheetKey } from '../syntax.js'
import { hi, LINES, LIMIT, parseRange, withSpan, type Area, type Axis } from './area.js'
import { placer, type Applied, type Geometry, type Place } from './geometry.js'
import { arraysOf, gridOf } from './grid.js'
import { relocated, shiftRefusal } from './arrays.js'
import { openedGroup, shiftFrozen, shiftLines, shiftMerges } from './layout.js'
import { offsetsStay, rewriteCell, rewriteCells, rewriteFormula, rewriteNames, sheetIndex, sheetNames, type Outcome, type Transform } from './rewrite.js'
import { copyTransform, deleteSpan, insertSpan } from './transform.js'
import { emptiedTable, tableShift } from '../tables/shift.js'
import { shiftConditional } from '../cond/structure.js'
import { madeFrom } from '../looks/sheet.js'
import { rewriteValidations, shiftValidations } from '../validation/structure.js'

type SpanFn = (lo: number, hi: number, at: number, n: number, max: number, cell: boolean) => [number, number] | null

/**
 * The transform of every reference to `sheet` along `axis`; `edge` gives the span of a reference into a table's data
 * that rows inserted into the table grow at its edge (`tables/shift.ts`).
 */
export function shiftTransform(sheet: string, axis: Axis, at: number, n: number, fn: SpanFn, edge?: (area: Area) => [number, number] | undefined): Transform {
  const key = sheetKey(sheet)
  return {
    ref({ sheet: s, area }): Outcome {
      if (sheetKey(s) !== key) return undefined
      const next = edge?.(area) ?? fn(area[axis].lo.i, area[axis].hi.i, at, n, LIMIT[axis], area.kind === 'cell')
      if (!next) return 'lost'
      return next[0] === area[axis].lo.i && next[1] === area[axis].hi.i ? undefined : { sheet: s, area: withSpan(area, axis, next[0], next[1]) }
    },
  }
}

/**
 * Inserts (`n` > 0) or deletes (`n` < 0) rows or columns at index `at` of `sheet`, with the sheet's layout.
 * Inserting refuses when an entry or a merged area would leave the sheet (Excel refuses for cells that are
 * not empty); the sheet's looks pushed off are dropped, as Excel drops formatting (`linesMoved`).
 */
export function shift(input: BookInput, sheet: string, axis: Axis, at: number, n: number): Applied {
  const emptied = axis === 'r' && n < 0 ? emptiedTable(input, sheet, at, -n) : null
  if (emptied) return keepFirstRow(input, sheet, at, -n, emptied)
  const tables = tableShift(input, sheet, axis, at, n)
  if (tables.refuse) return { ok: false, code: 'table', message: say(input.locale, tables.refuse) }
  const book = tables.before(input)
  const count = Math.abs(n)
  const ti = sheetIndex(book, sheet)
  const target = book.sheets[ti]
  const grid = gridOf(target)
  const line = axis === 'r' ? grid.row : grid.col
  const lines = LINES[axis]
  const refuse = (code: 'off_sheet' | 'array', m: EngineMessage): Applied => ({ ok: false, code, message: say(input.locale, m) })
  if (n > 0) {
    const off = line.findIndex((i) => i >= LIMIT[axis] - count)
    if (off >= 0) return refuse('off_sheet', { key: 'insertOffSheet', count, lines, sheet: target.name, cell: grid.addr[off] })
    // the cells of a legacy array's area past its top cell hold values in Excel, though not entries here
    const array = arraysOf(target).find(([, x]) => hi(x, axis) >= LIMIT[axis] - count)
    if (array) return refuse('off_sheet', { key: 'insertOffSheet', count, lines, sheet: target.name, cell: array[0] })
    if (tables.offSheet) return refuse('off_sheet', { key: 'insertOffSheet', count, lines, sheet: target.name, cell: tables.offSheet })
  }
  const cut = shiftRefusal(target, axis, at, n)
  if (cut) return refuse('array', cut)
  const merges = shiftMerges(target.merges, axis, at, n)
  if (merges === null) return refuse('off_sheet', { key: 'insertMergeOffSheet', count, lines, sheet: target.name })
  const geometry: Geometry = [{ kind: 'lines', s: ti, axis, at, n }]
  const t = shiftTransform(target.name, axis, at, count, n > 0 ? insertSpan : deleteSpan, tables.edge)
  const names = sheetNames(book)
  const sheets = book.sheets.map((s) => {
    const moving = s === target
    let cells: Record<string, CellInput> = {}
    let changed = moving
    if (moving) cells = placedCells(s, ti, geometry, t, names)
    else {
      cells = rewriteCells(s.cells, s.name, s.name, t, names)
      changed = cells !== s.cells
      if (changed) madeFrom([s.cells], cells)
    }
    const list = s.tables?.map((ti) => {
      const first = parseRange(ti.range)!.r1 + (ti.headers !== false ? 1 : 0)
      // If the first data row is deleted, take the calculated formula at the first surviving row.
      const dr = moving && axis === 'r' && n < 0 && first >= at && first < at + count ? at + count - first : 0
      const columns = ti.columns.map((c) => {
        if (c.formula === undefined) return c
        const source = dr ? rewriteFormula(c.formula, s.name, s.name, copyTransform(dr, 0), names) : c.formula
        const formula = rewriteFormula(source, s.name, s.name, t, names)
        return formula === c.formula ? c : { ...c, formula }
      })
      if (columns.every((c, i) => c === ti.columns[i])) return ti
      changed = true
      return { ...ti, columns }
    })
    const tableFields = list ? { tables: list } : {}
    const validations = moving ? shiftValidations(s.validations, s.name, axis, at, n, t, names) : rewriteValidations(s.validations, s.name, s.name, t, names)
    if (!moving) return changed || validations !== s.validations ? { ...s, cells, ...tableFields, ...(validations ? { validations } : {}) } : s
    // the new lines take the format of the line before them with the sheet's looks (`linesMoved`)
    madeFrom([s.cells], cells)
    const after = (axis === 'r' ? s.outline?.summaryBelow : s.outline?.summaryRight) !== false
    const layout = { ...s, cells, ...tableFields, [lines]: openedGroup(shiftLines(s[lines], axis, at, n), axis, at, n, after), merges, frozen: shiftFrozen(s.frozen, axis, at, n), validations }
    for (const k of ['columns', 'rows', 'merges', 'frozen', 'validations'] as const) if (layout[k] === undefined) delete layout[k]
    // lines and merged areas all pushed off or deleted leave no field, as a paste leaves none
    for (const k of ['columns', 'rows', 'merges'] as const) if (layout[k] && !Object.keys(layout[k]).length) delete layout[k]
    return layout
  })
  const out: BookInput = { ...book, sheets: shiftConditional(sheets, target.name, axis, at, n, t, names) }
  const nm = rewriteNames(book.names, offsetsStay(t, axis), names)
  if (nm) out.names = nm
  return { ok: true, book: tables.after(out), geometry }
}

/** The cells of sheet `ti` where the geometry takes them, their formulas rewritten by `t`. */
export function placedCells(sheet: SheetInput, ti: number, geometry: Geometry, t: Transform, names: Map<string, string>): Record<string, CellInput> {
  const [grid, place, p] = [gridOf(sheet), placer(geometry), { s: ti, r: 0, c: 0 } as Place]
  const cells: Record<string, CellInput> = {}
  for (let i = 0; i < grid.addr.length; i++) {
    const addr = grid.addr[i]
    p.s = ti
    p.r = grid.row[i]
    p.c = grid.col[i]
    if (!place(p)) continue
    const to = p.r === grid.row[i] && p.c === grid.col[i] ? addr : formatAddress(p.r, p.c)
    cells[to] = relocated(rewriteCell(grid.cell[i], sheet.name, sheet.name, t, names), addr, to)
  }
  return cells
}

/** Rows deleted over every data row of a table: its first data row stays, its cells in the table cleared; the rest go. */
function keepFirstRow(book: BookInput, sheet: string, at: number, count: number, cols: { c1: number; c2: number }): Applied {
  const r: Applied = count > 1 ? shift(book, sheet, 'r', at + 1, 1 - count) : { ok: true, book, geometry: [] }
  if (!r.ok) return r
  const i = sheetIndex(r.book, sheet)
  const sheets = r.book.sheets.map((s, j) => {
    if (j !== i) return s
    const cells = { ...s.cells }
    for (let c = cols.c1; c <= cols.c2; c++) {
      const addr = formatAddress(at, c)
      const cell = cells[addr]
      if (!cell) continue
      const { value: _, formula: __, array: ___, held: ____, ...look } = cell
      if (Object.keys(look).length) cells[addr] = look
      else delete cells[addr]
    }
    return { ...s, cells }
  })
  return { ok: true, book: { ...r.book, sheets }, geometry: [...r.geometry, { kind: 'clear', s: i, area: { r1: at, r2: at, ...cols } }] }
}
