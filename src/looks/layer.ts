/**
 * A sheet's looks held by rows (`SheetInput.looks`): each row has a pattern, the runs of its columns that show one look,
 * and rows next to each other with equal patterns are one block; equal patterns are held once however many rows use
 * them. A cell's look is found by its row's block and a binary search of the pattern. A layer never changes: an edit
 * gives a new one, which shares the patterns it leaves as they are.
 */
import type { NumberFormat } from '../contract.js'
import type { Look } from '../layout/lines.js'
import { MAX_COLS } from '../model.js'
import { bands, lowerBound } from '../sorted.js'
import type { Rect } from '../structure/area.js'

/** No look: a cell in a gap of a pattern. */
export const NONE = -1

/** A pattern: runs `[c1, c2, look]` of columns, in order, apart, two runs of one look never next to each other. */
export type Pattern = Int32Array

/** An empty pattern: a row without looks. */
export const EMPTY: Pattern = new Int32Array(0)

const same = (x: number) => x

const formatKey = (f: NumberFormat | undefined) => (f === undefined ? '-' : typeof f === 'number' ? `n${f}` : `s${f}`)

/** Looks by number: equal pairs of a style and a format share a number and one frozen object, `{ style, numberFormat }` with the fields given. */
export class LookTable {
  readonly list: Look[]
  private readonly ids: Map<string, number>

  constructor(base?: LookTable) {
    this.list = base ? base.list.slice() : []
    this.ids = base ? new Map(base.ids) : new Map()
  }

  /** The number of a look; NONE for neither field. */
  id(style: number | undefined, format: NumberFormat | undefined): number {
    if (style === undefined && format === undefined) return NONE
    const k = `${style ?? '-'}|${formatKey(format)}`
    let i = this.ids.get(k)
    if (i === undefined) {
      const look: Look = {}
      if (style !== undefined) look.style = style
      if (format !== undefined) look.numberFormat = format
      this.ids.set(k, (i = this.list.push(Object.freeze(look)) - 1))
    }
    return i
  }

  /** The number of a cell's or a rectangle's look, by its fields. */
  of(x: Look | undefined): number {
    return x ? this.id(x.style, x.numberFormat) : NONE
  }
}

/** A numeric digest of a pattern, which equal patterns share. */
function digest(p: Pattern): number {
  let h = 0x811c9dc5 ^ p.length
  for (let i = 0; i < p.length; i++) h = Math.imul(h ^ p[i], 0x01000193)
  return h >>> 0
}

function samePattern(a: Pattern, b: Pattern): boolean {
  if (a === b) return true
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

/** A pattern from runs that may touch and may be of no look: neighbours of one look joined, gaps dropped. */
export function patternOf(runs: ArrayLike<number>): Pattern {
  const out: number[] = []
  for (let i = 0; i < runs.length; i += 3) {
    const [c1, c2, id] = [runs[i], runs[i + 1], runs[i + 2]]
    if (id === NONE || c2 < c1) continue
    const n = out.length
    if (n && out[n - 1] === id && out[n - 2] === c1 - 1) out[n - 2] = c2
    else out.push(c1, c2, id)
  }
  return out.length ? Int32Array.from(out) : EMPTY
}

/** The spans of columns where two patterns each give one look (NONE for none), `[c1, c2, look in a, look in b]`, across the sheet. */
export function* spansOf(a: Pattern, b: Pattern): Generator<[number, number, number, number]> {
  let [i, j, c] = [0, 0, 0]
  while (c < MAX_COLS) {
    while (i < a.length && a[i + 1] < c) i += 3
    while (j < b.length && b[j + 1] < c) j += 3
    const inA = i < a.length && a[i] <= c
    const inB = j < b.length && b[j] <= c
    // the next column where either pattern changes
    const end = Math.min(inA ? a[i + 1] + 1 : i < a.length ? a[i] : MAX_COLS, inB ? b[j + 1] + 1 : j < b.length ? b[j] : MAX_COLS)
    yield [c, end - 1, inA ? a[i + 2] : NONE, inB ? b[j + 2] : NONE]
    c = end
  }
}

/** The look's number of the run of a pattern holding column `c`, NONE for none. */
export function runAt(p: Pattern, c: number): number {
  let [lo, hi] = [0, p.length / 3 - 1]
  while (lo <= hi) {
    const m = (lo + hi) >> 1
    if (p[m * 3] > c) hi = m - 1
    else if (p[m * 3 + 1] < c) lo = m + 1
    else return p[m * 3 + 2]
  }
  return NONE
}

/**
 * `top` over `base`: where a run of `top` lies its look (NONE clears), elsewhere `base`'s. `top`'s runs are in order and
 * apart, and may be of no look.
 */
function overlay(base: Pattern, top: ArrayLike<number>): Pattern {
  if (!top.length) return base
  const out: number[] = []
  let i = 0
  // the part of `base` before column `end`, from column `from` on
  const baseUpTo = (from: number, end: number) => {
    while (i < base.length && base[i + 1] < from) i += 3
    for (let k = i; k < base.length && base[k] < end; k += 3) out.push(Math.max(base[k], from), Math.min(base[k + 1], end - 1), base[k + 2])
  }
  let at = 0
  for (let j = 0; j < top.length; j += 3) {
    baseUpTo(at, top[j])
    out.push(top[j], top[j + 1], top[j + 2])
    at = top[j + 1] + 1
  }
  baseUpTo(at, MAX_COLS)
  return patternOf(out)
}

/** Patterns held once, found by their content. */
class Patterns {
  readonly list: Pattern[] = []
  private readonly byDigest = new Map<number, number[]>()
  private readonly known = new Map<Pattern, number>()

  /** The number of a pattern, an equal one's when there is one; NONE for an empty one. */
  of(p: Pattern): number {
    if (!p.length) return NONE
    const k = this.known.get(p)
    if (k !== undefined) return k
    const d = digest(p)
    const same = this.byDigest.get(d)
    for (const i of same ?? []) if (samePattern(this.list[i], p)) return i
    const i = this.list.push(p) - 1
    if (same) same.push(i)
    else this.byDigest.set(d, [i])
    this.known.set(p, i)
    return i
  }
}

/** Builds a layer from rows given in order: blocks of rows with one pattern, equal patterns held once. */
export class LayerBuilder {
  readonly looks: LookTable
  private readonly patterns = new Patterns()
  private readonly r1: number[] = []
  private readonly r2: number[] = []
  private readonly at: number[] = []

  constructor(looks?: LookTable) {
    this.looks = looks ?? new LookTable()
  }

  /** Rows `lo` to `hi` (after every row given before) with pattern `p`. */
  rows(lo: number, hi: number, p: Pattern): void {
    const k = this.patterns.of(p)
    if (k === NONE) return
    const n = this.at.length
    if (n && this.at[n - 1] === k && this.r2[n - 1] === lo - 1) this.r2[n - 1] = hi
    else {
      this.r1.push(lo)
      this.r2.push(hi)
      this.at.push(k)
    }
  }

  /** The last row given, -1 before any. */
  get last(): number {
    return this.r2.length ? this.r2[this.r2.length - 1] : -1
  }

  build(): Layer {
    return new Layer(this.looks, Int32Array.from(this.r1), Int32Array.from(this.r2), Int32Array.from(this.at), this.patterns.list)
  }
}

/** A piece of a layer: rows and columns of one look. */
export interface Piece extends Rect {
  look: number
}

export class Layer {
  constructor(
    readonly looks: LookTable,
    /** The blocks: their first and last rows, in order and apart, and their patterns by number. */
    readonly r1: Int32Array,
    readonly r2: Int32Array,
    readonly at: Int32Array,
    readonly patterns: readonly Pattern[],
  ) {}

  get empty(): boolean {
    return this.at.length === 0
  }

  get blocks(): number {
    return this.at.length
  }

  /** The block holding row `r`, or -1. */
  blockOf(r: number): number {
    const b = lowerBound(this.r2, r, same)
    return b < this.at.length && this.r1[b] <= r ? b : -1
  }

  /** The first block that ends at or after row `r`. */
  firstFrom(r: number): number {
    return lowerBound(this.r2, r, same)
  }

  /** The pattern of row `r`; an empty one for a row without looks. */
  row(r: number): Pattern {
    const b = this.blockOf(r)
    return b < 0 ? EMPTY : this.patterns[this.at[b]]
  }

  /** The look's number of a cell, NONE without one. */
  idAt(r: number, c: number): number {
    return runAt(this.row(r), c)
  }

  /** The look of a cell (a frozen object shared by the cells of that look), undefined without one. */
  lookAt(r: number, c: number): Look | undefined {
    const id = this.idAt(r, c)
    return id === NONE ? undefined : this.looks.list[id]
  }

  /** The pieces of the layer in rectangle `x`, row block by row block, each block's runs in order. */
  *pieces(x: Rect = FULL): Generator<Piece> {
    for (let b = this.firstFrom(x.r1); b < this.at.length && this.r1[b] <= x.r2; b++) {
      const [r1, r2] = [Math.max(this.r1[b], x.r1), Math.min(this.r2[b], x.r2)]
      const p = this.patterns[this.at[b]]
      for (let k = 0; k < p.length; k += 3) {
        if (p[k + 1] < x.c1) continue
        if (p[k] > x.c2) break
        yield { r1, r2, c1: Math.max(p[k], x.c1), c2: Math.min(p[k + 1], x.c2), look: p[k + 2] }
      }
    }
  }

  /** The looks its cells show, each once. */
  used(): Look[] {
    const ids = new Set<number>()
    for (const k of new Set(this.at)) for (let i = 2; i < this.patterns[k].length; i += 3) ids.add(this.patterns[k][i])
    return [...ids].map((id) => this.looks.list[id])
  }

  /** The rectangle the looks lie in, or null for none. */
  bounds(): Rect | null {
    if (this.empty) return null
    let [c1, c2] = [MAX_COLS, -1]
    for (const k of new Set(this.at)) {
      const p = this.patterns[k]
      c1 = Math.min(c1, p[0])
      c2 = Math.max(c2, p[p.length - 2])
    }
    return { r1: this.r1[0], r2: this.r2[this.r2.length - 1], c1, c2 }
  }
}

const FULL: Rect = { r1: 0, r2: Number.MAX_SAFE_INTEGER, c1: 0, c2: MAX_COLS - 1 }

/** An empty layer. */
export const emptyLayer = (): Layer => new LayerBuilder().build()

/** An edit of a layer: rows `lo` to `hi` get the runs `top` over what they have (`overlay`). */
export interface RowsEdit {
  lo: number
  hi: number
  top: ArrayLike<number>
}

/**
 * The layer with edits (in order of rows, apart) applied, its looks those of `looks` (the layer's own, or a table that
 * extends them); the layer itself when nothing changes.
 */
export function edited(layer: Layer, edits: RowsEdit[], looks: LookTable = layer.looks): Layer {
  if (!edits.length) return layer
  const out = new LayerBuilder(looks)
  // the rows where a block or an edit starts or ends, between them one old pattern and at most one edit
  const cuts = new Set<number>()
  for (let b = 0; b < layer.at.length; b++) cuts.add(layer.r1[b]).add(layer.r2[b] + 1)
  for (const e of edits) cuts.add(e.lo).add(e.hi + 1)
  const rows = bands(cuts)
  const done = new Map<string, Pattern>()
  let changed = false
  let [b, j] = [0, 0]
  for (const [lo, hi] of rows) {
    while (b < layer.at.length && layer.r2[b] < lo) b++
    while (j < edits.length && edits[j].hi < lo) j++
    const inBlock = b < layer.at.length && layer.r1[b] <= lo
    const p = inBlock ? layer.patterns[layer.at[b]] : EMPTY
    const e = j < edits.length && edits[j].lo <= lo ? edits[j] : undefined
    let next = p
    if (e) {
      const key = `${inBlock ? layer.at[b] : NONE}:${j}`
      next = done.get(key) ?? overlay(p, e.top)
      done.set(key, next)
      changed ||= !samePattern(next, p)
    }
    out.rows(lo, hi, next)
  }
  return changed || looks !== layer.looks ? out.build() : layer
}

/** Cells' new looks (NONE for none), by row and column, as edits of their rows. */
export function cellEdits(cells: Iterable<[r: number, c: number, look: number]>): RowsEdit[] {
  const byRow = new Map<number, number[][]>()
  for (const [r, c, look] of cells) {
    let row = byRow.get(r)
    if (!row) byRow.set(r, (row = []))
    row.push([c, look])
  }
  return [...byRow.keys()]
    .sort((a, b) => a - b)
    .map((r) => {
      // the last look given a cell counts
      const cols = new Map(byRow.get(r)!.map(([c, look]) => [c, look]))
      const top: number[] = []
      for (const c of [...cols.keys()].sort((a, b) => a - b)) top.push(c, c, cols.get(c)!)
      return { lo: r, hi: r, top }
    })
}

/** Whether two looks (of one table or two, none for one of them) are the same: their style and number format fields. */
export const sameLook = (x: Look | undefined, y: Look | undefined): boolean => x?.style === y?.style && x?.numberFormat === y?.numberFormat

/** Whether two layers give every cell the same look. */
export function sameLooks(a: Layer, b: Layer): boolean {
  const [x, y] = [[...a.pieces()], [...b.pieces()]]
  if (x.length !== y.length) return false
  return x.every((p, i) => {
    const q = y[i]
    return p.r1 === q.r1 && p.r2 === q.r2 && p.c1 === q.c1 && p.c2 === q.c2 && sameLook(a.looks.list[p.look], b.looks.list[q.look])
  })
}

/** A layer from cells given row by row, as a file lists them: a row's cells in order of columns; a cell out of that order is taken at the end, over the rest. */
export class CellsLayer {
  private readonly out = new LayerBuilder()
  private row = -1
  private runs: number[] = []
  private late: [number, number, number][] = []
  /** The cells of the current row given out of order, taken after its runs. */
  private again: [number, number, number][] = []

  get looks(): LookTable {
    return this.out.looks
  }

  /** The cell at row `r`, column `c` with look `look` (NONE: none). */
  add(r: number, c: number, look: number): void {
    if (look === NONE) return
    if (r !== this.row) {
      this.finish()
      this.row = r
    }
    const n = this.runs.length
    if (n && this.runs[n - 2] >= c) this.again.push([r, c, look])
    else if (n && this.runs[n - 1] === look && this.runs[n - 2] === c - 1) this.runs[n - 2] = c
    else this.runs.push(c, c, look)
  }

  private finish() {
    if (!this.runs.length) return
    const p = patternOf(this.runs)
    if (this.out.last < this.row) this.out.rows(this.row, this.row, p)
    // a row given again after later ones: its cells over what it has
    else for (let k = 0; k < p.length; k += 3) for (let c = p[k]; c <= p[k + 1]; c++) this.late.push([this.row, c, p[k + 2]])
    // the row's cells given out of order come after its runs, so that the last look given a cell counts
    for (const x of this.again) this.late.push(x)
    this.runs = []
    this.again = []
  }

  build(): Layer {
    this.finish()
    return edited(this.out.build(), cellEdits(this.late))
  }
}
