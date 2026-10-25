/**
 * The layout of a move or copy, as Excel changes it (checked in `test/excel/struct/styles.json`): whole rows
 * or columns take their entries (height or width, style, format) along; each pasted position shows the style
 * and format its source showed, spelled out on the cell where the rows and columns there would show another;
 * merged areas inside the source go along, a copy takes the part of one it cuts, and those the pasted cells
 * land on are gone. A move that cuts a merged area at its source, and a copy that lands on part of one, are
 * refused, as Excel refuses them.
 */
import type { BookInput, CellInput, ColumnInput, RowInput, SheetInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { lookOf, type Look } from '../layout/lines.js'
import { mergesOf, parseMerge } from '../layout/merges.js'
import { cut, entryAt, lines, slice, toMap, type Line } from '../layout/spans.js'
import { say, type EngineMessage } from '../locale/messages.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { bands } from '../sorted.js'
import { cellCount, contains, cornersText, hasCell, hi, intersect, lo, meets, shiftRect, type Axis, type Rect } from './area.js'
import { arrayRefusal } from './arrays.js'
import type { Applied } from './geometry.js'
import { cellsIn, gridOf } from './grid.js'
import { sheetOf } from './rewrite.js'
import { cellEdits, edited, emptyLayer, LookTable, NONE, rectEdits, type Layer, type Piece } from '../looks/layer.js'
import { layerOf, setLayer } from '../looks/sheet.js'
import { sheetKey } from '../syntax.js'

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

/** Excel's refusal of the paste, or null: a merged area it cuts or lands on part of (`merged`), else an array formula's area (`array`). */
export function pasteRefusal(book: BookInput, p: Paste): Exclude<Applied, { ok: true }> | null {
  const refused = mergeRefusal(book, p)
  if (refused) return { ok: false, code: 'merged', message: say(book.locale, refused) }
  const part = arrayRefusal(book, p)
  return part && { ok: false, code: 'array', message: say(book.locale, part) }
}

/** Why Excel refuses the paste, or null: a move cutting a merged area at its source, a copy landing on part of one. */
function mergeRefusal(book: BookInput, p: Paste): EngineMessage | null {
  const [sheet, rect] = p.move ? [p.from, p.src] : [p.to, p.area]
  const m = mergesOf(sheetOf(book, sheet)).find((x) => meets(x, rect) && !contains(rect, x))
  if (!m) return null
  return { key: p.move ? 'moveCutsMerge' : 'pasteOnMerge', sheet, area: cornersText(m) }
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
  if (!src.length && !dst.length) return [[a, b]]
  const at = new Set([a, b + 1])
  bounds(src, a, b, 0, at)
  bounds(dst, a, b, d, at)
  return bands([...at].filter((x) => x >= a && x <= b + 1))
}

/** The bounds of the lines `ls` (shifted back by `d`) that can fall in [a, b + 1], added to `at`: a fill pastes one row per tile, so the lines in order are halved to the first, not all walked. */
function bounds(ls: Line<unknown>[], a: number, b: number, d: number, at: Set<number>) {
  let [i, j] = [0, ls.length]
  while (i < j) {
    const m = (i + j) >> 1
    if (ls[m][1] + 1 - d < a) i = m + 1
    else j = m
  }
  for (; i < ls.length && ls[i][0] - d <= b + 1; i++) at.add(ls[i][0] - d).add(ls[i][1] + 1 - d)
}

/**
 * The lines of a paste's old source sheet and of its target, the source's cells with the look each shows there, the
 * looks of its cells without an entry and the target's after the paste (`looks/move.ts`).
 */
interface Source {
  sr: Line<RowInput>[]
  sc: Line<ColumnInput>[]
  tr: Line<RowInput>[]
  tc: Line<ColumnInput>[]
  cells: { row: number; col: number; look: Look }[]
  looks: Layer | undefined
}

/** The cells of `src` on the old sheet `from`, found once for all the offsets of a paste (a fill pastes one row thousands of times). */
function sourceOf(from: SheetInput, to: SheetInput, src: Rect): Source {
  const out: Source = { sr: lines(from.rows, 'r'), sc: lines(from.columns, 'c'), tr: lines(to.rows, 'r'), tc: lines(to.columns, 'c'), cells: [], looks: layerOf(from) }
  const g = gridOf(from)
  for (const i of cellsIn(g, src)) {
    const [row, col, cell] = [g.row[i], g.col[i], g.cell[i]]
    const s = lineLook(out.sr, out.sc, row, col)
    out.cells.push({ row, col, look: { style: cell.style ?? s.style, numberFormat: cell.numberFormat ?? s.numberFormat } })
  }
  return out
}

/** The fields `cell` takes to show look `s` where the rows and columns show `t`: a style with nothing set (`plain()`) or `General` where `s` has none. */
function spelled(cell: Look, s: Look, t: Look, plain: () => number): Look {
  const style = (cell.style ?? t.style) !== s.style ? { style: s.style ?? plain() } : {}
  const format = (cell.numberFormat ?? t.numberFormat) !== s.numberFormat ? { numberFormat: s.numberFormat ?? 'General' } : {}
  return { ...style, ...format }
}

/**
 * Each position of `src` on the old sheet, pasted at (dr, dc) of `to`, shows what it showed: where the rows and columns
 * of `to` would give another style or format, the cell spells out the source's (`spelled`). A cell of an entry spells it
 * in its entry; the others, which show the source's looks pasted on the target's, by pieces of those looks, into `laid`
 * (rectangles of the target with looks of `looks`), never cell by cell: a paste of whole formatted columns covers millions.
 * The cells of entries the pieces lay over go to `under`, as `[row, column]` of the target.
 */
function keepLooks(to: SheetInput, src: Rect, source: Source, dr: number, dc: number, plain: () => number, laid: Piece[], under: [number, number][], looks: LookTable) {
  const { sr, sc, tr, tc } = source
  // the parts of the source whose lines show another look than the target's there
  const apart: [Rect, Look, Look][] = []
  for (const [ra, rb] of segments(sr, tr, src.r1, src.r2, dr))
    for (const [ca, cb] of segments(sc, tc, src.c1, src.c2, dc)) {
      const [s, t] = [lineLook(sr, sc, ra, ca), lineLook(tr, tc, ra + dr, ca + dc)]
      if (s.style !== t.style || s.numberFormat !== t.numberFormat) apart.push([{ r1: ra, r2: rb, c1: ca, c2: cb }, s, t])
    }
  for (const { row, col, look } of source.cells) {
    spend(1)
    const addr = formatAddress(row + dr, col + dc)
    const cell = to.cells[addr] ?? {}
    const more = spelled(cell, look, lineLook(tr, tc, row + dr, col + dc), plain)
    if (Object.keys(more).length) to.cells[addr] = { ...cell, ...more }
    if (apart.some(([x]) => hasCell(x, row, col))) under.push([row + dr, col + dc])
  }
  for (const [x, s, t] of apart)
    // a piece of the source's looks shows its look over the lines', pasted on the target's as it is
    for (const p of source.looks?.pieces(x, true) ?? [{ ...x, look: NONE }]) {
      spend(1)
      const look = p.look === NONE ? {} : source.looks!.looks.list[p.look]
      const more = spelled(look, { style: look.style ?? s.style, numberFormat: look.numberFormat ?? s.numberFormat }, t, plain)
      if (Object.keys(more).length) laid.push({ r1: p.r1 + dr, r2: p.r2 + dr, c1: p.c1 + dc, c2: p.c2 + dc, look: looks.of({ ...look, ...more }) })
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
    const x = intersect(m, p.src)
    return x && (x.r1 < x.r2 || x.c1 < x.c2) && (!p.move || contains(p.src, m)) ? [x] : []
  })
  if (p.move) from.merges = (from.merges ?? []).filter((m) => !contains(p.src, parseMerge(m)!))
  const kept = (to.merges ?? []).filter((m) => !meets(parseMerge(m)!, p.area))
  to.merges = [...kept, ...p.offsets.flatMap(([dr, dc]) => carried.map((m) => cornersText(shiftRect(m, dr, dc))))]
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
  const base = layerOf(to)
  const looks = new LookTable(base?.looks)
  const [laid, under]: [Piece[], [number, number][]] = [[], []]
  for (const [dr, dc] of p.offsets) keepLooks(to, p.src, source, dr, dc, plain, laid, under, looks)
  if (!laid.length) return
  // the cells of entries keep the looks under them, as when each position was written as an entry
  const layer = edited(edited(base ?? emptyLayer(), rectEdits(laid), looks), cellEdits(under.map(([r, c]): [number, number, number] => [r, c, base?.idAt(r, c) ?? NONE])), looks)
  setLayer(to, layer)
  if (laid.reduce((n, x) => n + cellCount(x), 0) > under.length) loose?.add(sheetKey(to.name))
}

/** The sheets of the operation running whose looks a paste laid where no entry is (`withLooseLooks`). */
let loose: Set<string> | null = null

/**
 * Operation `f` and the keys of the sheets whose looks a paste laid where no entry is: those positions were entries of a
 * look alone, which the engine took into the looks as the sheet's last field (`withLayer`) once the outermost operation
 * ended (a fill's copies are operations inside it), and so stays its order.
 */
export function withLooseLooks<T>(f: () => T): [T, Set<string>] {
  const outer = loose
  const mine = (loose = new Set())
  try {
    return [f(), mine]
  } finally {
    loose = outer
    if (outer) for (const k of mine) outer.add(k)
  }
}
