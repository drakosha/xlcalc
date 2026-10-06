/**
 * Inserting and deleting rows and columns, with references adjusted as Excel adjusts them (checked in
 * `test/excel/struct/rows.json` and `edge.json`): a range grows when rows are inserted inside it and
 * shrinks when some of its rows are deleted, a deleted reference becomes `#REF!`, absolute references
 * move too. At the sheet's edge a range ending on the last row keeps ending there, a single cell pushed
 * past it is `#REF!`, and a range over every row is not changed by rows inserted or deleted.
 */
import type { BookInput, CellInput, CellRef, RestructureResult } from '../contract.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { LIMIT, withSpan, type Axis } from './area.js'
import { formatBefore, shiftLines, shiftMerges } from './layout.js'
import { rewriteCell, rewriteNames, sheetNames, type Outcome, type Transform } from './rewrite.js'

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

/** The transform of every reference to `sheet` along `axis`. */
function shiftTransform(sheet: string, axis: Axis, at: number, n: number, fn: SpanFn): Transform {
  const key = sheetKey(sheet)
  return {
    ref({ sheet: s, area }): Outcome {
      if (sheetKey(s) !== key) return undefined
      const next = fn(area[axis].lo.i, area[axis].hi.i, at, n, LIMIT[axis], area.kind === 'cell')
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
export function shift(book: BookInput, sheet: string, axis: Axis, at: number, n: number): RestructureResult {
  const count = Math.abs(n)
  const max = LIMIT[axis]
  const pos = (addr: string) => {
    const p = parseAddress(addr)
    return axis === 'r' ? p.row : p.col
  }
  const place = (addr: string, i: number) => {
    const p = parseAddress(addr)
    return axis === 'r' ? formatAddress(i, p.col) : formatAddress(p.row, i)
  }
  const target = book.sheets.find((s) => sheetKey(s.name) === sheetKey(sheet))!
  const lines = axis === 'r' ? 'rows' : 'columns'
  if (n > 0) {
    const off = Object.keys(target.cells).find((a) => pos(a) >= max - count)
    if (off) return { ok: false, code: 'off_sheet', message: `inserting ${count} ${lines} would push ${target.name}!${off} off the sheet` }
  }
  const merges = shiftMerges(target.merges, axis, at, n)
  if (merges === null) return { ok: false, code: 'off_sheet', message: `inserting ${count} ${lines} would push a merged area of ${target.name} off the sheet` }
  const t = shiftTransform(target.name, axis, at, count, n > 0 ? insertSpan : deleteSpan)
  const names = sheetNames(book)
  const sheets = book.sheets.map((s) => {
    const moving = s === target
    const cells: Record<string, CellInput> = {}
    let changed = moving
    for (const [addr, cell] of Object.entries(s.cells)) {
      const next = rewriteCell(cell, s.name, s.name, t, names)
      changed ||= next !== cell
      if (!moving) cells[addr] = next
      else {
        const i = pos(addr)
        if (n > 0) cells[i >= at ? place(addr, i + count) : addr] = next
        else if (i < at) cells[addr] = next
        else if (i >= at + count) cells[place(addr, i - count)] = next
      }
    }
    if (!moving) return changed ? { ...s, cells } : s
    // the new lines take the format of the line before them
    if (n > 0) Object.assign(cells, formatBefore(s, axis, at, n))
    const layout = { ...s, cells, [lines]: shiftLines(s[lines], axis, at, n), merges }
    for (const k of ['columns', 'rows', 'merges'] as const) if (layout[k] === undefined) delete layout[k]
    return layout
  })
  const out: BookInput = { ...book, sheets }
  const nm = rewriteNames(book.names, t, names)
  if (nm) out.names = nm
  const key = sheetKey(target.name)
  const mapCell = (ref: CellRef): CellRef | null => {
    if (sheetKey(ref.sheet) !== key) return ref
    const i = pos(ref.cell)
    if (n > 0) return i >= at ? { sheet: ref.sheet, cell: place(ref.cell, i + count) } : ref
    if (i < at) return ref
    return i >= at + count ? { sheet: ref.sheet, cell: place(ref.cell, i - count) } : null
  }
  return { ok: true, book: out, mapCell }
}
