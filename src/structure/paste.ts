/**
 * The layout of a move or copy, as Excel changes it (checked in `test/excel/struct/styles.json`): whole rows
 * or columns take their entries (height or width, style, format) along; each pasted position shows the style
 * and format its source showed, spelled out on the cell where the rows and columns there would show another;
 * merged areas inside the source go along, a copy takes the part of one it cuts, and those the pasted cells
 * land on are gone. A move that cuts a merged area at its source, and a copy that lands on part of one, are
 * refused, as Excel refuses them.
 */
import type { BookInput, ColumnInput, RowInput, SheetInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { lookOf, type Look } from '../layout/lines.js'
import { mergeText, parseMerge } from '../layout/merges.js'
import { cut, entryAt, lines, slice, toMap, type Line } from '../layout/spans.js'
import type { Message } from '../locale/messages.js'
import { formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { contains, hi, lo, meets, type Axis, type Rect } from './area.js'

export interface Paste {
  from: string
  src: Rect
  to: string
  /** Where the pasted cells land on `to`. */
  area: Rect
  /** The offset of each pasted copy of `src` (one for a move). */
  offsets: [number, number][]
  move: boolean
}

const sheetOf = (book: BookInput, name: string) => book.sheets.find((s) => sheetKey(s.name) === sheetKey(name))!
const mergesOf = (s: SheetInput) => (s.merges ?? []).map((m) => parseMerge(m)!)
const shifted = (x: Rect, dr: number, dc: number): Rect => ({ r1: x.r1 + dr, r2: x.r2 + dr, c1: x.c1 + dc, c2: x.c2 + dc })

/** Why Excel refuses the paste, or null: a move cutting a merged area at its source, a copy landing on part of one. */
export function mergeRefusal(book: BookInput, p: Paste): Message | null {
  const [sheet, rect] = p.move ? [p.from, p.src] : [p.to, p.area]
  const m = mergesOf(sheetOf(book, sheet)).find((x) => meets(x, rect) && !contains(rect, x))
  if (!m) return null
  return { key: p.move ? 'moveCutsMerge' : 'pasteOnMerge', sheet, area: mergeText(m) }
}

/** A whole line's entries moved or copied: `taken` from the old source, pasted at each offset of `d`. */
function pasteLines<T>(from: SheetInput, to: SheetInput, taken: Line<T>[], key: 'rows' | 'columns', axis: Axis, span: [number, number], ds: number[], move: boolean) {
  const map = (s: SheetInput) => s[key] as Record<string, T> | undefined
  if (move) Object.assign(from, { [key]: toMap(cut(lines(map(from), axis), ...span), axis) })
  let dest = lines(map(to), axis)
  for (const d of ds) dest = [...cut(dest, span[0] + d, span[1] + d), ...slice(taken, span[0], span[1], d)]
  Object.assign(to, { [key]: toMap(dest, axis) })
}

/** The style and format the rows and columns give at a position (`lookOf`). */
const lineLook = (rows: Line<RowInput>[], cols: Line<ColumnInput>[], r: number, c: number): Look => lookOf(undefined, entryAt(rows, r), entryAt(cols, c))

/** The bounds of segments of [a, b] where the source's lines and the destination's (shifted by `d`) do not change. */
function segments(src: Line<unknown>[], dst: Line<unknown>[], a: number, b: number, d: number): [number, number][] {
  const at = new Set([a, b + 1])
  for (const [x, y] of src) at.add(x).add(y + 1)
  for (const [x, y] of dst) at.add(x - d).add(y + 1 - d)
  const bs = [...at].filter((x) => x >= a && x <= b + 1).sort((x, y) => x - y)
  return bs.slice(0, -1).map((x, i) => [x, bs[i + 1] - 1])
}

/** The lines of a paste's old source sheet and of its target, and the source's cells with the look each shows there. */
interface Source {
  sr: Line<RowInput>[]
  sc: Line<ColumnInput>[]
  tr: Line<RowInput>[]
  tc: Line<ColumnInput>[]
  cells: { row: number; col: number; look: Look }[]
  /** The positions of `cells`, as `row * MAX_COLS + col`. */
  own: Set<number>
}

/** The cells of `src` on the old sheet `from`, found once for all the offsets of a paste (a fill pastes one row thousands of times). */
function sourceOf(from: SheetInput, to: SheetInput, src: Rect): Source {
  const out: Source = { sr: lines(from.rows, 'r'), sc: lines(from.columns, 'c'), tr: lines(to.rows, 'r'), tc: lines(to.columns, 'c'), cells: [], own: new Set() }
  for (const addr of Object.keys(from.cells)) {
    spend(1)
    const p = parseAddress(addr)
    if (p.row < src.r1 || p.row > src.r2 || p.col < src.c1 || p.col > src.c2) continue
    const cell = from.cells[addr]
    const s = lineLook(out.sr, out.sc, p.row, p.col)
    out.own.add(p.row * MAX_COLS + p.col)
    out.cells.push({ row: p.row, col: p.col, look: { style: cell.style ?? s.style, numberFormat: cell.numberFormat ?? s.numberFormat } })
  }
  return out
}

/**
 * Each position of `src` on the old sheet, pasted at (dr, dc) of `to`, shows what it showed: where the
 * rows and columns of `to` would give another style or format, the cell spells out the source's (a style with
 * nothing set, `plain()`, or `General` where the source had none).
 */
function keepLooks(to: SheetInput, src: Rect, source: Source, dr: number, dc: number, plain: () => number) {
  const { sr, sc, tr, tc } = source
  const spell = (addr: string, s: Look, t: Look) => {
    const cell = to.cells[addr] ?? {}
    const style = (cell.style ?? t.style) !== s.style ? { style: s.style ?? plain() } : {}
    const format = (cell.numberFormat ?? t.numberFormat) !== s.numberFormat ? { numberFormat: s.numberFormat ?? 'General' } : {}
    if (Object.keys(style).length || Object.keys(format).length) to.cells[addr] = { ...cell, ...style, ...format }
  }
  for (const { row, col, look } of source.cells) {
    spend(1)
    spell(formatAddress(row + dr, col + dc), look, lineLook(tr, tc, row + dr, col + dc))
  }
  for (const [ra, rb] of segments(sr, tr, src.r1, src.r2, dr))
    for (const [ca, cb] of segments(sc, tc, src.c1, src.c2, dc)) {
      const [s, t] = [lineLook(sr, sc, ra, ca), lineLook(tr, tc, ra + dr, ca + dc)]
      if (s.style === t.style && s.numberFormat === t.numberFormat) continue
      for (let r = ra; r <= rb; r++)
        for (let c = ca; c <= cb; c++) {
          spend(1)
          if (!source.own.has(r * MAX_COLS + c)) spell(formatAddress(r + dr, c + dc), s, t)
        }
    }
}

/**
 * The layout of `out` after the paste (its `from` and `to` sheets are its own objects, with the pasted cells):
 * whole lines' entries, merged areas, the looks of the pasted positions; a plain style is added to the book's
 * styles when a position needs one and there is none.
 */
export function pasteLayout(old: BookInput, out: BookInput, p: Paste) {
  const [oldFrom, from, to] = [sheetOf(old, p.from), sheetOf(out, p.from), sheetOf(out, p.to)]
  const whole: Record<Axis, boolean> = { r: p.src.c1 === 0 && p.src.c2 === MAX_COLS - 1, c: p.src.r1 === 0 && p.src.r2 === MAX_ROWS - 1 }
  if (whole.r) pasteLines(from, to, lines<RowInput>(oldFrom.rows, 'r'), 'rows', 'r', [lo(p.src, 'r'), hi(p.src, 'r')], p.offsets.map(([dr]) => dr), p.move)
  if (whole.c) pasteLines(from, to, lines<ColumnInput>(oldFrom.columns, 'c'), 'columns', 'c', [lo(p.src, 'c'), hi(p.src, 'c')], p.offsets.map(([, dc]) => dc), p.move)
  // merged areas: inside the source they go along (a copy takes the part it cuts), those landed on are gone
  const carried = mergesOf(oldFrom).flatMap((m) => {
    const x: Rect = { r1: Math.max(m.r1, p.src.r1), r2: Math.min(m.r2, p.src.r2), c1: Math.max(m.c1, p.src.c1), c2: Math.min(m.c2, p.src.c2) }
    return x.r1 <= x.r2 && x.c1 <= x.c2 && (x.r1 < x.r2 || x.c1 < x.c2) && (!p.move || contains(p.src, m)) ? [x] : []
  })
  if (p.move) from.merges = (from.merges ?? []).filter((m) => !contains(p.src, parseMerge(m)!))
  const kept = (to.merges ?? []).filter((m) => !meets(parseMerge(m)!, p.area))
  to.merges = [...kept, ...p.offsets.flatMap(([dr, dc]) => carried.map((m) => mergeText(shifted(m, dr, dc))))]
  // a copy leaves its source sheet as it was
  for (const s of p.move ? [from, to] : [to]) for (const k of ['rows', 'columns', 'merges'] as const) if (s[k] && !Object.keys(s[k]).length) delete s[k]
  let plainStyle: number | undefined
  const plain = () => {
    if (plainStyle !== undefined) return plainStyle
    const styles = out.styles ?? []
    plainStyle = styles.findIndex((s) => Object.keys(s).length === 0)
    if (plainStyle < 0) {
      out.styles = [...styles, {}]
      plainStyle = styles.length
    }
    return plainStyle
  }
  const source = sourceOf(oldFrom, to, p.src)
  for (const [dr, dc] of p.offsets) keepLooks(to, p.src, source, dr, dc, plain)
}
