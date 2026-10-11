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
import { InputError } from '../input.js'
import { sheetLines, type SheetLines } from '../layout/lines.js'
import { lines as spansOf, patchLines, spanKey } from '../layout/spans.js'
import { formatAddress, parseAddress } from '../model.js'
import type { Rect } from '../structure/area.js'
import type { CheckedPatch, Look, StyleChange, Target } from './patch.js'

/** The cells a rectangle may give a look to: as many as a whole column has. Past them, whole columns or rows cost what their cells hold. */
export const MAX_CELLS = 1_048_576

/** A key of a value independent of the order of its fields. */
function keyOf(v: unknown): string {
  if (v === null || typeof v !== 'object') return JSON.stringify(v)
  const o = v as Record<string, unknown>
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${keyOf(o[k])}`)
    .join(',')}}`
}

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

/** The book's styles, found by their fields: an equal style's number, or a new one added at the end. */
export class StyleTable {
  private keys: Map<string, number> | null = null
  private list: CellStyle[]
  added = false

  constructor(private readonly base: CellStyle[] | undefined) {
    this.list = base ?? []
  }

  /** The style of a number; the plain style for none. */
  at(i: number | undefined): CellStyle {
    return i === undefined ? {} : this.list[i]
  }

  isPlain(i: number | undefined): boolean {
    return i === undefined || keyOf(this.list[i]) === '{}'
  }

  indexOf(style: CellStyle): number {
    if (!this.keys) {
      this.keys = new Map()
      this.list.forEach((s, i) => this.keys!.has(keyOf(s)) || this.keys!.set(keyOf(s), i))
    }
    const k = keyOf(style)
    let i = this.keys.get(k)
    if (i === undefined) {
      if (!this.added) this.list = [...this.list]
      this.added = true
      i = this.list.push(style) - 1
      this.keys.set(k, i)
    }
    return i
  }

  /** The new table, when styles were added. */
  get table(): CellStyle[] | undefined {
    return this.added ? this.list : undefined
  }
}

/** A plan of looks: the cells' new inputs, the sheets with new lines or rules, the lines whose cells may show otherwise, the style table when it grew. */
export interface Plan {
  edits: CellReplace[]
  sheets: Map<number, SheetInput>
  touched: LayoutEdit[]
  styles?: CellStyle[]
}

type Line = RowInput & ColumnInput

/** The patches' work: the cells and sheets as the patches before left them, over the session's. */
export class Draft {
  private readonly cells = new Map<string, CellInput | null>()
  readonly sheets = new Map<number, SheetInput>()
  readonly touched: LayoutEdit[] = []
  readonly table: StyleTable

  constructor(readonly s: EngineState) {
    this.table = new StyleTable(s.input.styles)
  }

  sheet(i: number): SheetInput {
    return this.sheets.get(i) ?? this.s.input.sheets[i]
  }

  input(i: number, addr: string): CellInput | undefined {
    const k = `${i}!${addr}`
    return this.cells.has(k) ? (this.cells.get(k) ?? undefined) : this.s.input.sheets[i].cells[addr]
  }

  set(i: number, addr: string, input: CellInput | null) {
    this.cells.set(`${i}!${addr}`, input)
  }

  /** The sheet's lines as the draft holds them. */
  lines(i: number): SheetLines {
    const sh = this.sheet(i)
    return sh === this.s.input.sheets[i] ? sheetLines(sh, this.s) : sheetLines({ cells: {}, rows: sh.rows, columns: sh.columns })
  }

  /** The cells of a rectangle that hold an input, as `[row, column, address]`. */
  *stored(i: number, x: Rect): Generator<[number, number, string]> {
    const area = (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)
    if (area <= 65_536) {
      for (let r = x.r1; r <= x.r2; r++)
        for (let c = x.c1; c <= x.c2; c++) {
          const a = formatAddress(r, c)
          if (this.input(i, a)) yield [r, c, a]
        }
      return
    }
    const seen = new Set<string>()
    const keys = [...Object.keys(this.s.input.sheets[i].cells), ...[...this.cells.keys()].filter((k) => k.startsWith(`${i}!`)).map((k) => k.slice(k.indexOf('!') + 1))]
    for (const a of keys) {
      if (seen.has(a)) continue
      seen.add(a)
      const { row, col } = parseAddress(a)
      if (row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2 && this.input(i, a)) yield [row, col, a]
    }
  }

  /** The plan: every cell the looks went to, changed or not, as a call's step keeps the cells it was given. */
  plan(): Plan {
    const edits: CellReplace[] = []
    for (const [k, input] of this.cells) {
      const at = k.indexOf('!')
      edits.push({ sheet: this.s.sheetNames[Number(k.slice(0, at))], cell: k.slice(at + 1), input })
    }
    return { edits, sheets: this.sheets, touched: this.touched, ...(this.table.table ? { styles: this.table.table } : {}) }
  }
}

/** The fields of a cell's content. */
const CONTENT = ['value', 'formula', 'array'] as const

/**
 * A cell's input with a look over it: `before` and `after` are its sheet's lines before and after the look's change of
 * them. `own`: the cell is of a rectangle, which gives every cell the look; a cell of formatted lines takes a style or a
 * format of its own only where its lines would not show it.
 */
function lookOver(d: Draft, input: CellInput | undefined, r: number, c: number, look: Look, before: SheetLines, after: SheetLines, own: boolean): CellInput | null {
  const out: Record<string, unknown> = { ...input }
  if (look.content) for (const k of CONTENT) delete out[k]
  const st = look.style
  if (st) {
    const shown = after.row(r)?.style ?? after.column(c)?.style
    const want = 'set' in st ? st.set : 'merge' in st ? d.table.indexOf(patchStyle(d.table.at(input?.style ?? before.row(r)?.style ?? before.column(c)?.style), st.merge)) : 'reset' in st && !d.table.isPlain(shown) ? d.table.indexOf({}) : undefined
    if (want === undefined) delete out.style
    else if (input?.style !== undefined || own || want !== shown) out.style = want
  }
  const ft = look.format
  if (ft) {
    const shown = after.row(r)?.format ?? after.column(c)?.format
    const want = 'set' in ft ? ft.set : 'reset' in ft && shown !== undefined && shown !== 0 && shown !== 'General' ? 'General' : undefined
    if (want === undefined) delete out.format
    else if (input?.format !== undefined || own || want !== shown) out.format = want
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

/** A rectangle's cells with a look: every cell when the look gives one, or resets a look its lines show; else the cells that hold an input. */
function lookRect(d: Draft, t: Target, look: Look) {
  const x = t.rect
  const ls = d.lines(t.i)
  const sh = d.sheet(t.i)
  const reset = 'reset' in (look.style ?? {}) || 'reset' in (look.format ?? {})
  const looked = (map: Record<string, Line> | undefined, axis: 'r' | 'c', lo: number, hi: number) => spansOf(map, axis).some(([a, b, e]) => b >= lo && a <= hi && (e.style !== undefined || e.format !== undefined))
  const lined = reset && (looked(sh.rows, 'r', x.r1, x.r2) || looked(sh.columns, 'c', x.c1, x.c2))
  if (gives(look) || lined) {
    const area = (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)
    if (area > MAX_CELLS) throw new InputError('range', `a rectangle of ${area} cells: a look goes to ${MAX_CELLS} cells of a rectangle at most, or to whole columns or rows (B:D, 2:9)`)
    for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) d.set(t.i, formatAddress(r, c), lookOver(d, d.input(t.i, formatAddress(r, c)), r, c, look, ls, ls, true))
    return
  }
  for (const [r, c, a] of [...d.stored(t.i, x)]) d.set(t.i, a, lookOver(d, d.input(t.i, a), r, c, look, ls, ls, true))
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
  const field = axis === 'r' ? 'rows' : 'columns'
  const [lo, hi] = axis === 'r' ? [t.rect.r1, t.rect.r2] : [t.rect.c1, t.rect.c2]
  const sh = d.sheet(t.i)
  const before = d.lines(t.i)
  if (look.style || look.format) {
    let map = sh[field] as Record<string, Line> | undefined
    for (const [a, b, e] of spans(map, axis, lo, hi)) {
      const patch: LinePatch<Line> = {}
      if (look.style) patch.style = lineStyle(d, e?.style, look.style)
      if (look.format) patch.format = 'set' in look.format ? look.format.set : null
      map = patchLines(map, axis, a, b, patch)
    }
    const next: SheetInput = { ...sh, [field]: map }
    if (!Object.keys(map ?? {}).length) delete next[field]
    d.sheets.set(t.i, next)
    if (look.format) d.touched.push({ sheet: d.s.sheetNames[t.i], [field]: spanKey(lo, hi, axis), set: {} } as LayoutEdit)
  }
  const after = d.lines(t.i)
  const cells = new Map<string, [number, number]>()
  for (const [r, c, a] of d.stored(t.i, t.rect)) cells.set(a, [r, c])
  // where the lines cross lines of the other direction that show a style or a format the look changes
  const st = look.style && !('remove' in look.style)
  const ft = look.format && !('remove' in look.format)
  if (st || ft) {
    const crossing = spansOf((axis === 'r' ? sh.columns : sh.rows) as Record<string, Line> | undefined, axis === 'r' ? 'c' : 'r').filter(([, , e]) => (st && e.style !== undefined) || (ft && e.format !== undefined))
    const count = crossing.reduce((n, [a, b]) => n + (b - a + 1) * (hi - lo + 1), 0)
    if (count > MAX_CELLS) throw new InputError('range', `the lines cross formatted lines in ${count} cells, past ${MAX_CELLS}`)
    for (const [a, b] of crossing)
      for (let k = a; k <= b; k++)
        for (let j = lo; j <= hi; j++) {
          const [r, c] = axis === 'r' ? [j, k] : [k, j]
          cells.set(formatAddress(r, c), [r, c])
        }
  }
  for (const [a, [r, c]] of cells) d.set(t.i, a, lookOver(d, d.input(t.i, a), r, c, look, before, after, false))
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
