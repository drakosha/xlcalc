/**
 * The look of cells and ranges changed (`Session.patchCells`, and `clearCells` for formats): a plan of the cells' new
 * inputs, the sheets' new lines and the book's style table with the styles it takes, an equal style found rather than
 * added. A style by its fields goes over the style a cell shows, its own or else its row's or column's, as Excel formats
 * a cell over its look. Whole rows and columns are formatted as Excel formats them (`docs/edits.md`): the lines' own
 * styles and formats, the cells of theirs with one of their own, and the cells where they cross a line of the other
 * direction with one, which Excel gives the two merged; an empty cell of a rectangle gets a cell of its own.
 */
import type { CellInput, CellReplace, CellStyle, ColumnInput, LayoutEdit, LinePatch, RowInput, SheetInput, StylePatch } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { sheetLines, type SheetLines } from '../layout/lines.js'
import { lines as spansOf, patchLines, spanKey } from '../layout/spans.js'
import { formatAddress, parseAddress } from '../model.js'
import { cellCount, hasCell, LINES, type Rect } from '../structure/area.js'
import { cellsIn, gridOf, type Grid } from '../structure/grid.js'
import { StyleTable } from '../style-table.js'
import { edited, emptyLayer, LookTable, NONE, type Layer, type RowsEdit } from '../looks/layer.js'
import { layerOf, lookAlone, hasLook } from '../looks/sheet.js'
import { inputAt } from './book.js'
import { CONTENT } from './cells.js'
import type { CheckedPatch, Look, StyleChange, Target } from './patch.js'

/** A rectangle of at most so many cells is read cell by cell. */
const SMALL = 65_536

/** The cells a rectangle may give a look to: as many as a whole column has. Past them, whole columns or rows cost what their cells hold. */
export const MAX_CELLS = 1_048_576

/** An object's fields with a patch over them: null removes a field; undefined when nothing is left. */
function over<T extends object>(base: T | undefined, patch: { [K in keyof T]?: T[K] | null } | null | undefined): T | undefined {
  if (patch === null) return undefined
  if (patch === undefined) return base
  const out: Record<string, unknown> = { ...base }
  for (const [k, v] of Object.entries(patch)) {
    if (v === null) delete out[k]
    else if (v !== undefined) out[k] = v
  }
  for (const k in out) if (out[k] !== undefined) return out as T
  return undefined
}

/** A style with a patch over it, field by field: the font, the border's sides and the alignment one field at a time, the fill whole. */
export function patchStyle(base: CellStyle, p: StylePatch): CellStyle {
  const out: CellStyle = {}
  const font = over(base.font, p.font)
  const fill = p.fill === undefined ? base.fill : (p.fill ?? undefined)
  const border = over(base.border, p.border)
  const alignment = over(base.alignment, p.alignment)
  if (font) out.font = font
  if (fill) out.fill = fill
  if (border) out.border = border
  if (alignment) out.alignment = alignment
  return out
}

/**
 * A plan of looks: the cells' new inputs, the sheets' looks changed by rectangles (laid before the cells), the sheets
 * with new lines or rules, the lines whose cells may show otherwise, the style table when it grew.
 */
export interface Plan {
  edits: CellReplace[]
  layers?: Map<number, Layer>
  sheets: Map<number, SheetInput>
  touched: LayoutEdit[]
  styles?: CellStyle[]
}

type Line = RowInput & ColumnInput

/** The patches' work: the cells and sheets as the patches before left them, over the session's. */
export class Draft {
  private readonly cells = new Map<string, CellInput | null>()
  /** The addresses of `cells` by sheet, and the session's cells of a sheet with their places, read once for the draft. */
  private readonly drafted = new Map<number, Set<string>>()
  private readonly grids = new Map<number, Grid>()
  /** The sheets' looks as rectangles of the patches before changed them (`relook`), over the session's. */
  private readonly layers = new Map<number, Layer>()
  readonly sheets = new Map<number, SheetInput>()
  readonly touched: LayoutEdit[] = []
  readonly table: StyleTable

  constructor(readonly s: EngineState) {
    this.table = new StyleTable(s.input.styles)
  }

  /** The sheet's looks as the draft holds them. */
  layer(i: number): Layer | undefined {
    return this.layers.get(i) ?? layerOf(this.s.input.sheets[i])
  }

  /** The sheet's looks with edits of rows (`RowsEdit`, in order) of looks of `looks`, a table over theirs. */
  relook(i: number, edits: RowsEdit[], looks: LookTable) {
    if (edits.length) this.layers.set(i, edited(this.layer(i) ?? emptyLayer(), edits, looks))
  }

  sheet(i: number): SheetInput {
    return this.sheets.get(i) ?? this.s.input.sheets[i]
  }

  /** A cell's input as the draft holds it, a look of the sheet's looks as an input of a look alone. */
  input(i: number, addr: string): CellInput | undefined {
    const k = `${i}!${addr}`
    if (this.cells.has(k)) return this.cells.get(k) ?? undefined
    const layer = this.layers.get(i)
    if (!layer) return inputAt(this.s, i, addr)
    const { row, col } = parseAddress(addr)
    return this.s.input.sheets[i].cells[addr] ?? layer.lookAt(row, col)
  }

  set(i: number, addr: string, input: CellInput | null) {
    this.cells.set(`${i}!${addr}`, input)
    let own = this.drafted.get(i)
    if (!own) this.drafted.set(i, (own = new Set()))
    own.add(addr)
  }

  /** The sheet's lines as the draft holds them. */
  lines(i: number): SheetLines {
    const sh = this.sheet(i)
    return sh === this.s.input.sheets[i] ? sheetLines(sh, this.s) : sheetLines({ cells: {}, rows: sh.rows, columns: sh.columns })
  }

  /** Whether the draft holds a cell's input. */
  has(i: number, addr: string): boolean {
    return !!this.drafted.get(i)?.has(addr)
  }

  /** The cells of a rectangle the draft holds an input of (or none, `null`), as `[row, column, address]`. */
  *held(i: number, x: Rect): Generator<[number, number, string]> {
    for (const a of this.drafted.get(i) ?? []) {
      const { row, col } = parseAddress(a)
      if (hasCell(x, row, col)) yield [row, col, a]
    }
  }

  /**
   * The cells of a rectangle that hold an input or a look of the sheet's looks, as `[row, column, address]`; past
   * SMALL cells without those of the looks when `looks` is false (`byPieces` reads them by pieces).
   */
  *stored(i: number, x: Rect, looks = true): Generator<[number, number, string]> {
    const area = cellCount(x)
    if (area <= SMALL) {
      for (let r = x.r1; r <= x.r2; r++)
        for (let c = x.c1; c <= x.c2; c++) {
          const a = formatAddress(r, c)
          if (this.input(i, a)) yield [r, c, a]
        }
      return
    }
    const sheet = this.s.input.sheets[i]
    let g = this.grids.get(i)
    if (!g) this.grids.set(i, (g = gridOf(sheet)))
    for (const k of cellsIn(g, x)) if (this.input(i, g.addr[k])) yield [g.row[k], g.col[k], g.addr[k]]
    // the cells the draft added
    const drafted = this.drafted.get(i)
    for (const a of drafted ?? []) {
      if (a in sheet.cells) continue
      const { row, col } = parseAddress(a)
      if (hasCell(x, row, col) && this.input(i, a)) yield [row, col, a]
    }
    // the cells of a look of the sheet's looks alone
    if (!looks) return
    for (const p of this.layer(i)?.pieces(x) ?? [])
      for (let r = p.r1; r <= p.r2; r++)
        for (let c = p.c1; c <= p.c2; c++) {
          const a = formatAddress(r, c)
          if (!(a in sheet.cells) && !drafted?.has(a)) yield [r, c, a]
        }
  }

  /** The plan: every cell the looks went to, changed or not, as a call's step keeps the cells it was given. */
  plan(): Plan {
    const edits: CellReplace[] = []
    for (const [k, input] of this.cells) {
      const at = k.indexOf('!')
      edits.push({ sheet: this.s.sheetNames[Number(k.slice(0, at))], cell: k.slice(at + 1), input })
    }
    return { edits, ...(this.layers.size ? { layers: this.layers } : {}), sheets: this.sheets, touched: this.touched, ...(this.table.table ? { styles: this.table.table } : {}) }
  }
}

/**
 * A cell's input with a look over it: `before` and `after` are its sheet's lines before and after the look's change of
 * them. `own`: the cell is of a rectangle, which gives every cell the look; a cell of formatted lines takes a style or a
 * format of its own only where its lines would not show it.
 */
function lookOver(table: StyleTable, input: CellInput | undefined, r: number, c: number, look: Look, before: SheetLines, after: SheetLines, own: boolean): CellInput | null {
  const out: Record<string, unknown> = { ...input }
  if (look.content) for (const k of CONTENT) delete out[k]
  const st = look.style
  if (st) {
    const shown = after.row(r)?.style ?? after.column(c)?.style
    const want = 'set' in st ? st.set : 'merge' in st ? table.indexOf(patchStyle(table.at(input?.style ?? before.row(r)?.style ?? before.column(c)?.style), st.merge)) : 'reset' in st && !table.isPlain(shown) ? table.indexOf({}) : undefined
    if (want === undefined) delete out.style
    else if (input?.style !== undefined || own || want !== shown) out.style = want
  }
  const ft = look.format
  if (ft) {
    const shown = after.row(r)?.numberFormat ?? after.column(c)?.numberFormat
    const want = 'set' in ft ? ft.set : 'reset' in ft && shown !== undefined && shown !== 0 && shown !== 'General' ? 'General' : undefined
    if (want === undefined) delete out.numberFormat
    else if (input?.numberFormat !== undefined || own || want !== shown) out.numberFormat = want
  }
  for (const k of ['meta', 'note', 'link'] as const) {
    if (look[k] === null) delete out[k]
    else if (look[k] !== undefined) out[k] = look[k]
  }
  for (const k in out) if (out[k] !== undefined) return out as CellInput
  return null
}

/** Whether a look gives an empty cell an input: a field set, or a reset where the cell's lines show a look. */
const gives = (look: Look) => ('set' in (look.style ?? {}) || 'merge' in (look.style ?? {}) || 'set' in (look.format ?? {}) || [look.meta, look.note, look.link].some((x) => x !== undefined && x !== null))

/**
 * A rectangle's cells with a look: every cell when the look gives one, or resets a look its lines show; else the cells
 * that hold an input. Past SMALL cells, those of the sheet's looks (every one: those of no look too) by pieces, the rest
 * cell by cell.
 */
function lookRect(d: Draft, t: Target, look: Look) {
  const x = t.rect
  const ls = d.lines(t.i)
  const sh = d.sheet(t.i)
  const reset = 'reset' in (look.style ?? {}) || 'reset' in (look.format ?? {})
  const looked = (map: Record<string, Line> | undefined, axis: 'r' | 'c', lo: number, hi: number) => spansOf(map, axis).some(([a, b, e]) => b >= lo && a <= hi && hasLook(e))
  const every = gives(look) || (reset && (looked(sh.rows, 'r', x.r1, x.r2) || looked(sh.columns, 'c', x.c1, x.c2)))
  const area = cellCount(x)
  if (every && area > MAX_CELLS) throw new InputError('range', `a rectangle of ${area} cells: formatting applies to at most ${MAX_CELLS} cells of a rectangle, or to whole columns or rows (B:D, 2:9)`)
  if (area > SMALL) {
    const pieces = byPieces(d, t.i, x, look, ls, ls, true, every, [sh])
    const cells = new Map<string, [number, number]>()
    for (const [r, c, a] of [...d.stored(t.i, x, false), ...(every ? d.held(t.i, x) : []), ...pieces.cells]) cells.set(a, [r, c])
    for (const [a, [r, c]] of cells) d.set(t.i, a, lookOver(d.table, d.input(t.i, a), r, c, look, ls, ls, true))
    d.relook(t.i, pieces.edits, pieces.looks)
    return
  }
  if (every) {
    for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) d.set(t.i, formatAddress(r, c), lookOver(d.table, d.input(t.i, formatAddress(r, c)), r, c, look, ls, ls, true))
    return
  }
  for (const [r, c, a] of [...d.stored(t.i, x)]) d.set(t.i, a, lookOver(d.table, d.input(t.i, a), r, c, look, ls, ls, true))
}

/** The style a line takes from a change: one of the table, fields over its own, or none. */
function lineStyle(d: Draft, style: number | undefined, st: StyleChange): number | null {
  if ('set' in st) return st.set
  if (!('merge' in st)) return null
  const next = patchStyle(d.table.at(style), st.merge)
  return Object.keys(next).length ? d.table.indexOf(next) : null
}

/** The spans of lines [lo, hi] with the entry of each, gaps without one. */
function spans(map: Record<string, Line> | undefined, axis: 'r' | 'c', lo: number, hi: number): [number, number, Line | undefined][] {
  const out: [number, number, Line | undefined][] = []
  let next = lo
  for (const [a, b, e] of spansOf(map, axis)) {
    if (b < lo || a > hi) continue
    if (a > next) out.push([next, a - 1, undefined])
    out.push([Math.max(a, lo), Math.min(b, hi), e])
    next = b + 1
  }
  if (next <= hi) out.push([next, hi, undefined])
  return out
}

/**
 * Whole rows or columns with a look: their own entries, the cells of theirs that hold an input, and the cells where they
 * cross a line of the other direction whose style or format the look changes the cell's view of.
 */
function lookLines(d: Draft, t: Target, look: Look) {
  const axis = t.lines!
  const field = LINES[axis]
  const [lo, hi] = axis === 'r' ? [t.rect.r1, t.rect.r2] : [t.rect.c1, t.rect.c2]
  const sh = d.sheet(t.i)
  const before = d.lines(t.i)
  if (look.style || look.format) {
    let map = sh[field] as Record<string, Line> | undefined
    for (const [a, b, e] of spans(map, axis, lo, hi)) {
      const patch: LinePatch<Line> = {}
      if (look.style) patch.style = lineStyle(d, e?.style, look.style)
      if (look.format) patch.numberFormat = 'set' in look.format ? look.format.set : null
      map = patchLines(map, axis, a, b, patch)
    }
    const next: SheetInput = { ...sh, [field]: map }
    if (!Object.keys(map ?? {}).length) delete next[field]
    d.sheets.set(t.i, next)
    if (look.format) d.touched.push({ sheet: d.s.sheetNames[t.i], [field]: spanKey(lo, hi, axis), set: {} } as LayoutEdit)
  }
  const after = d.lines(t.i)
  const cells = new Map<string, [number, number]>()
  const big = cellCount(t.rect) > SMALL
  for (const [r, c, a] of d.stored(t.i, t.rect, !big)) cells.set(a, [r, c])
  const pieces = big ? byPieces(d, t.i, t.rect, look, before, after, false, false, [sh, d.sheet(t.i)]) : undefined
  for (const [r, c, a] of pieces?.cells ?? []) cells.set(a, [r, c])
  // where the lines cross lines of the other direction that show a style or a format the look changes
  const st = look.style && !('remove' in look.style)
  const ft = look.format && !('remove' in look.format)
  if (st || ft) {
    const crossing = spansOf((axis === 'r' ? sh.columns : sh.rows) as Record<string, Line> | undefined, axis === 'r' ? 'c' : 'r').filter(([, , e]) => (st && e.style !== undefined) || (ft && e.numberFormat !== undefined))
    const count = crossing.reduce((n, [a, b]) => n + (b - a + 1) * (hi - lo + 1), 0)
    if (count > MAX_CELLS) throw new InputError('range', `the lines cross formatted lines in ${count} cells, more than ${MAX_CELLS}`)
    for (const [a, b] of crossing)
      for (let k = a; k <= b; k++)
        for (let j = lo; j <= hi; j++) {
          const [r, c] = axis === 'r' ? [j, k] : [k, j]
          cells.set(formatAddress(r, c), [r, c])
        }
  }
  for (const [a, [r, c]] of cells) d.set(t.i, a, lookOver(d.table, d.input(t.i, a), r, c, look, before, after, false))
  if (pieces) d.relook(t.i, pieces.edits, pieces.looks)
}

/** The rows (`r`) or columns (`c`) where an entry of the lines of `sheets` starts or ends, in order. */
const lineCuts = (sheets: SheetInput[], axis: 'r' | 'c') => [...new Set(sheets.flatMap((x) => spansOf((axis === 'r' ? x.rows : x.columns) as Record<string, Line> | undefined, axis).flatMap(([a, b]) => [a, b + 1])))].sort((a, b) => a - b)

/** The segments of [lo, hi] between the cuts `at` (in order). */
function segments(at: number[], lo: number, hi: number): [number, number][] {
  const out: [number, number][] = []
  let from = lo
  for (const k of at) {
    if (k <= from) continue
    if (k > hi) break
    out.push([from, k - 1])
    from = k
  }
  out.push([from, hi])
  return out
}

/**
 * The cells of rectangle `x` of the sheet's looks (`all`: and those of no look) with a look over them, as edits of the
 * draft's looks: a piece of one look under the same rows and columns, before and after (`sheets`), changes each of its
 * cells alike, so it is tried once, and a piece the look leaves as it is costs nothing per cell. The cells of a piece
 * that takes more than a look (a note, a link), without an entry, go cell by cell (`cells`); so do the cells of an entry
 * and those the draft holds, whose inputs the caller writes over these looks.
 */
function byPieces(d: Draft, i: number, x: Rect, look: Look, before: SheetLines, after: SheetLines, own: boolean, all: boolean, sheets: SheetInput[]): { edits: RowsEdit[]; looks: LookTable; cells: [number, number, string][] } {
  const layer = d.layer(i)
  const looks = new LookTable(layer?.looks)
  const [rowCuts, colCuts] = [lineCuts(sheets, 'r'), lineCuts(sheets, 'c')]
  // the bands of rows of one pattern of the looks and the same rows' entries
  const at = new Set([x.r1, x.r2 + 1, ...rowCuts.filter((k) => k > x.r1 && k <= x.r2)])
  for (let b = layer?.firstFrom(x.r1) ?? 0; layer && b < layer.blocks && layer.r1[b] <= x.r2; b++) for (const k of [layer.r1[b], layer.r2[b] + 1]) if (k > x.r1 && k <= x.r2) at.add(k)
  const rows = [...at].sort((a, b) => a - b)
  // a band's runs to lay and its spans of columns to write cell by cell, by its pattern and its rows' entries
  const ids = new Map<object | undefined, number>()
  const id = (o: object | undefined) => ids.get(o) ?? ids.set(o, ids.size).get(o)!
  const memo = new Map<string, { top: number[]; loose: [number, number][] }>()
  const edits: RowsEdit[] = []
  const cells: [number, number, string][] = []
  const entries = d.s.input.sheets[i].cells
  for (let k = 0; k + 1 < rows.length; k++) {
    const [lo, hi] = [rows[k], rows[k + 1] - 1]
    const p = layer?.row(lo)
    const key = `${id(p)}|${id(before.row(lo))}|${id(after.row(lo))}`
    let band = memo.get(key)
    if (!band) {
      band = { top: [], loose: [] }
      memo.set(key, band)
      // the runs of the pattern across the rectangle, with the gaps between them when `all`
      const runs: [number, number, number][] = []
      let c = x.c1
      for (let j = 0; p && j < p.length; j += 3) {
        if (p[j + 1] < x.c1) continue
        if (p[j] > x.c2) break
        const [a, b] = [Math.max(p[j], x.c1), Math.min(p[j + 1], x.c2)]
        if (all && a > c) runs.push([c, a - 1, NONE])
        runs.push([a, b, p[j + 2]])
        c = b + 1
      }
      if (all && c <= x.c2) runs.push([c, x.c2, NONE])
      for (const [a, b, was] of runs)
        for (const [c1, c2] of segments(colCuts, a, b)) {
          const next = lookOver(d.table, was === NONE ? undefined : layer!.looks.list[was], lo, c1, look, before, after, own)
          if (next && !lookAlone(next)) band.loose.push([c1, c2])
          else if (looks.of(next ?? undefined) !== was) band.top.push(c1, c2, looks.of(next ?? undefined))
        }
    }
    if (band.top.length) edits.push({ lo, hi, top: band.top })
    for (const [c1, c2] of band.loose)
      for (let r = lo; r <= hi; r++)
        for (let c = c1; c <= c2; c++) {
          const a = formatAddress(r, c)
          if (!(a in entries) && !d.has(i, a)) cells.push([r, c, a])
        }
  }
  return { edits, looks, cells }
}

/** A target with a look, in the draft. */
export function lookAt(d: Draft, t: Target, look: Look) {
  if (t.lines) lookLines(d, t, look)
  else lookRect(d, t, look)
}

/** The plan of patches (checked), each over what the ones before it left. */
export function patchPlan(s: EngineState, patches: CheckedPatch[]): Plan {
  const d = new Draft(s)
  for (const p of patches) lookAt(d, p, p.look)
  return d.plan()
}
