/**
 * Inserting and deleting rows and columns, with references adjusted as Excel adjusts them (checked in
 * `test/excel/struct/rows.json` and `edge.json`): a range grows when rows are inserted inside it and
 * shrinks when some of its rows are deleted, a deleted reference becomes `#REF!`, absolute references
 * move too. At the sheet's edge a range ending on the last row keeps ending there, a single cell pushed
 * past it is `#REF!`, and a range over every row is not changed by rows inserted or deleted.
 */
import type { BookInput, CellInput } from '../contract.js'
import { formatAddress } from '../model.js'
import { say, type Message } from '../locale/messages.js'
import { sheetKey } from '../syntax.js'
import { LIMIT, parseRange, withSpan, type Area, type Axis } from './area.js'
import { copyTransform } from './copy.js'
import { placer, type Applied, type Geometry, type Place } from './geometry.js'
import { gridOf } from './grid.js'
import { relocated, shiftRefusal } from './arrays.js'
import { formatBefore, openedGroup, shiftFrozen, shiftLines, shiftMerges } from './layout.js'
import { offsetsStay, rewriteCell, rewriteFormula, rewriteNames, sheetIndex, sheetNames, type Outcome, type Transform } from './rewrite.js'
import { emptiedTable, tableShift } from '../tables/shift.js'
import { shiftConditional } from '../cond/structure.js'
import { rewriteValidations, shiftValidations } from '../validation/structure.js'

/** One axis of a reference after `n` rows (columns) are inserted before index `at`; null when it leaves the sheet. */
export function insertSpan(lo: number, hi: number, at: number, n: number, max: number, cell: boolean): [number, number] | null {
  if (!cell && lo === 0 && hi === max - 1) return [lo, hi]
  const a = lo >= at ? lo + n : lo
  const b = hi >= at ? hi + n : hi
  if (a > max - 1 || (cell && b > max - 1)) return null
  return [a, Math.min(b, max - 1)]
}

/** One axis of a reference after `n` rows (columns) from index `at` are deleted; null when nothing of it is left. */
export function deleteSpan(lo: number, hi: number, at: number, n: number, max: number, cell: boolean): [number, number] | null {
  const end = at + n
  // an index in the deleted rows goes to `gone`: the first row after them, or the last before them
  const map = (i: number, gone: number) => (i < at ? i : i >= end ? i - n : gone)
  if (lo >= at && hi < end) return null
  if (cell) return [map(lo, lo), map(lo, lo)]
  const a = map(lo, at)
  // a range ending on the sheet's last row keeps ending there
  const b = hi === max - 1 ? hi : map(hi, at - 1)
  return a > b ? null : [a, b]
}

type SpanFn = (lo: number, hi: number, at: number, n: number, max: number, cell: boolean) => [number, number] | null

/**
 * The transform of every reference to `sheet` along `axis`; `edge` gives the span of a reference into a table's data
 * that rows inserted into the table grow at its edge (`tables/shift.ts`).
 */
function shiftTransform(sheet: string, axis: Axis, at: number, n: number, fn: SpanFn, edge?: (area: Area) => [number, number] | undefined): Transform {
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
 * Inserting refuses when a cell or a merged area would leave the sheet (Excel refuses for cells that are
 * not empty).
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
  const lines = axis === 'r' ? 'rows' : 'columns'
  const refuse = (code: 'off_sheet' | 'array', m: Message): Applied => ({ ok: false, code, message: say(input.locale, m) })
  if (n > 0) {
    const off = line.findIndex((i) => i >= LIMIT[axis] - count)
    if (off >= 0) return refuse('off_sheet', { key: 'insertOffSheet', count, lines, sheet: target.name, cell: grid.addr[off] })
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
    const cells: Record<string, CellInput> = {}
    let changed = moving
    if (moving) {
      // the cells where the geometry takes them, their formulas rewritten
      const place = placer(geometry)
      const p: Place = { s: ti, r: 0, c: 0 }
      for (let i = 0; i < grid.addr.length; i++) {
        const addr = grid.addr[i]
        p.s = ti
        p.r = grid.row[i]
        p.c = grid.col[i]
        if (!place(p)) continue
        const to = p.r === grid.row[i] && p.c === grid.col[i] ? addr : formatAddress(p.r, p.c)
        cells[to] = relocated(rewriteCell(grid.cell[i], s.name, s.name, t, names), addr, to)
      }
    } else
      for (const addr of Object.keys(s.cells)) {
        const cell = s.cells[addr]
        const next = rewriteCell(cell, s.name, s.name, t, names)
        changed ||= next !== cell
        cells[addr] = next
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
    // the new lines take the format of the line before them
    if (n > 0) Object.assign(cells, formatBefore(s, axis, at, n))
    const after = (axis === 'r' ? s.outline?.summaryBelow : s.outline?.summaryRight) !== false
    const layout = { ...s, cells, ...tableFields, [lines]: openedGroup(shiftLines(s[lines], axis, at, n), axis, at, n, after), merges, frozen: shiftFrozen(s.frozen, axis, at, n), validations }
    for (const k of ['columns', 'rows', 'merges', 'frozen', 'validations'] as const) if (layout[k] === undefined) delete layout[k]
    return layout
  })
  const out: BookInput = { ...book, sheets: shiftConditional(sheets, target.name, axis, at, n, t, names) }
  const nm = rewriteNames(book.names, offsetsStay(t), names)
  if (nm) out.names = nm
  return { ok: true, book: tables.after(out), geometry }
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
      const { value: _, formula: __, array: ___, ...look } = cell
      if (Object.keys(look).length) cells[addr] = look
      else delete cells[addr]
    }
    return { ...s, cells }
  })
  return { ok: true, book: { ...r.book, sheets }, geometry: [...r.geometry, { kind: 'clear', s: i, area: { r1: at, r2: at, ...cols } }] }
}
