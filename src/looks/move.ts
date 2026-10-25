/**
 * A layer moved as a structure operation moves cells (`structure/geometry.ts`), by its blocks and patterns, never cell by
 * cell: lines inserted (blank, or with the look of the line before them, as an operation gives new lines) and deleted,
 * rectangles cleared, cut and pasted. Then the layers of a book's sheets carried by a geometry, and the cells whose look
 * differs between two layers, which a session lists after an operation.
 */
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { LIMIT, shiftRect, type Axis, type Rect } from '../structure/area.js'
import type { Geometry } from '../structure/geometry.js'
import { edited, EMPTY, LayerBuilder, LookTable, NONE, patternOf, runAt, sameLook, spansOf, type Layer, type Pattern, type RowsEdit } from './layer.js'

/** A pattern's runs with `n` columns inserted before `at` (or `-n` deleted from `at`); `before`: the new columns take the look of column `at` − 1. */
function shiftedRuns(p: Pattern, at: number, n: number, before: boolean): Pattern {
  const out: number[] = []
  const push = (c1: number, c2: number, look: number) => {
    if (c1 < MAX_COLS) out.push(c1, Math.min(c2, MAX_COLS - 1), look)
  }
  const d = -n
  for (let k = 0; k < p.length; k += 3) {
    const [c1, c2, look] = [p[k], p[k + 1], p[k + 2]]
    if (c2 < at) push(c1, c2, look)
    else if (n > 0) {
      if (c1 < at) push(c1, at - 1, look)
      push(Math.max(c1, at) + n, c2 + n, look)
    } else if (c1 >= at + d) push(c1 - d, c2 - d, look)
    else {
      if (c1 < at) push(c1, at - 1, look)
      if (c2 >= at + d) push(at, c2 - d, look)
    }
  }
  const look = before && n > 0 && at > 0 ? runAt(p, at - 1) : NONE
  if (look !== NONE) push(at, at + n - 1, look)
  const runs: number[][] = []
  for (let k = 0; k < out.length; k += 3) runs.push(out.slice(k, k + 3))
  return patternOf(runs.sort((x, y) => x[0] - y[0]).flat())
}

/**
 * The layer with `n` lines inserted before index `at` along `axis` (`n` > 0) or `-n` lines from `at` deleted, the lines
 * after moved, those pushed past the sheet's end gone; `before`: the new lines take the look of the line before them.
 */
export function shifted(layer: Layer, axis: Axis, at: number, n: number, before = false): Layer {
  if (layer.empty) return layer
  const out = new LayerBuilder(layer.looks)
  if (axis === 'c') {
    const memo = new Map<Pattern, Pattern>()
    for (let b = 0; b < layer.blocks; b++) {
      const p = layer.patterns[layer.at[b]]
      let q = memo.get(p)
      if (!q) memo.set(p, (q = shiftedRuns(p, at, n, before)))
      out.rows(layer.r1[b], layer.r2[b], q)
    }
    return out.build()
  }
  const put = (lo: number, hi: number, p: Pattern) => {
    if (lo < MAX_ROWS) out.rows(lo, Math.min(hi, MAX_ROWS - 1), p)
  }
  const d = -n
  let placed = n < 0
  const place = () => {
    if (placed) return
    placed = true
    if (before && at > 0) put(at, at + n - 1, layer.row(at - 1))
  }
  for (let b = 0; b < layer.blocks; b++) {
    const [r1, r2, p] = [layer.r1[b], layer.r2[b], layer.patterns[layer.at[b]]]
    if (r2 < at) put(r1, r2, p)
    else if (n > 0) {
      if (r1 < at) put(r1, at - 1, p)
      place()
      put(Math.max(r1, at) + n, r2 + n, p)
    } else if (r1 >= at + d) put(r1 - d, r2 - d, p)
    else {
      if (r1 < at) put(r1, at - 1, p)
      if (r2 >= at + d) put(at, r2 - d, p)
    }
  }
  place()
  return out.build()
}

/** Rectangles as edits of rows that clear them: each band of rows with the columns the rectangles over it cover. */
export function clearing(xs: Rect[]): RowsEdit[] {
  const cuts = [...new Set(xs.flatMap((x) => [x.r1, x.r2 + 1]))].sort((a, b) => a - b)
  const out: RowsEdit[] = []
  for (let k = 0; k + 1 < cuts.length; k++) {
    const [lo, hi] = [cuts[k], cuts[k + 1] - 1]
    const cols = xs.filter((x) => x.r1 <= lo && x.r2 >= hi).map((x): [number, number] => [x.c1, x.c2]).sort((a, b) => a[0] - b[0])
    const top: number[] = []
    for (const [c1, c2] of cols) {
      const n = top.length
      if (n && top[n - 2] >= c1 - 1) top[n - 2] = Math.max(top[n - 2], c2)
      else top.push(c1, c2, NONE)
    }
    if (top.length) out.push({ lo, hi, top })
  }
  return out
}

/** The cells of `src` of layer `from` cut and pasted `dr` rows and `dc` columns away on layer `to` (the same or another), as a range move: the place they land on cleared, the rest of the sheet kept. */
export function movedLooks(from: Layer | undefined, src: Rect, to: Layer | undefined, dr: number, dc: number, same: boolean): [Layer | undefined, Layer | undefined] {
  const dst = shiftRect(src, dr, dc)
  const inside = { r1: Math.max(dst.r1, 0), r2: Math.min(dst.r2, LIMIT.r - 1), c1: Math.max(dst.c1, 0), c2: Math.min(dst.c2, LIMIT.c - 1) }
  const left = from && edited(from, clearing([src]))
  const base = same ? left : to
  const looks = new LookTable(base?.looks)
  // the pieces of the source, where they land, by rows in order (the source's blocks are)
  const edits: RowsEdit[] = []
  for (const x of from?.pieces(src) ?? []) {
    const [r1, r2] = [Math.max(x.r1 + dr, inside.r1), Math.min(x.r2 + dr, inside.r2)]
    const [c1, c2] = [Math.max(x.c1 + dc, inside.c1), Math.min(x.c2 + dc, inside.c2)]
    if (r1 > r2 || c1 > c2) continue
    const look = looks.of(from!.looks.list[x.look])
    const last = edits.at(-1)
    if (last && last.lo === r1 && last.hi === r2) (last.top as number[]).push(c1, c2, look)
    else edits.push({ lo: r1, hi: r2, top: [c1, c2, look] })
  }
  const landed = edited(base ? edited(base, clearing([inside])) : new LayerBuilder().build(), edits, looks)
  return same ? [landed, landed] : [left, landed]
}

/**
 * Layer `to` with the looks of `src` of layer `from` laid on each tile of a copy, which cover `area` from its top-left
 * cell (`tiles`), the area cleared first: as `copyRange` lays its cells, by rectangles. A source whose rows show one
 * pattern across it is laid once over every tile's rows. Undefined when neither layer has looks.
 */
export function pastedLooks(from: Layer | undefined, src: Rect, to: Layer | undefined, area: Rect): Layer | undefined {
  const base = to && edited(to, clearing([area]))
  const looks = new LookTable(base?.looks)
  // the source's bands of rows of one pattern, each with its runs across the source
  const bands: { lo: number; hi: number; runs: number[] }[] = []
  for (const x of from?.pieces(src) ?? []) {
    const last = bands.at(-1)
    const id = looks.of(from!.looks.list[x.look])
    if (last && last.lo === x.r1) last.runs.push(x.c1, x.c2, id)
    else bands.push({ lo: x.r1, hi: x.r2, runs: [x.c1, x.c2, id] })
  }
  if (!bands.length) return base
  // the tiles by rows, each row of tiles with the runs of every band across its columns
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  const dcs: number[] = []
  for (let c = area.c1; c <= area.c2; c += w) dcs.push(c - src.c1)
  const tops = bands.map((b) => dcs.flatMap((dc) => b.runs.map((v, k) => (k % 3 === 2 ? v : v + dc))))
  const whole = bands.length === 1 && bands[0].lo === src.r1 && bands[0].hi === src.r2
  const edits: RowsEdit[] = []
  if (whole) edits.push({ lo: area.r1, hi: area.r2, top: tops[0] })
  else for (let r = area.r1; r <= area.r2; r += h) for (let k = 0; k < bands.length; k++) edits.push({ lo: bands[k].lo + r - src.r1, hi: bands[k].hi + r - src.r1, top: tops[k] })
  return edited(base ?? new LayerBuilder().build(), edits, looks)
}

/** The layers of a book's sheets (by index) carried by a geometry: each cell's look where the geometry takes the cell, none where no cell comes. */
export function carried(layers: (Layer | undefined)[], g: Geometry): (Layer | undefined)[] {
  const out = layers.slice()
  for (const m of g)
    switch (m.kind) {
      case 'lines':
        out[m.s] = out[m.s] && shifted(out[m.s]!, m.axis, m.at, m.n)
        break
      case 'clear':
        out[m.s] = out[m.s] && edited(out[m.s]!, clearing([m.area]))
        break
      case 'range': {
        const [a, b] = movedLooks(out[m.from], m.src, out[m.to], m.dr, m.dc, m.from === m.to)
        out[m.from] = a
        out[m.to] = b
        break
      }
      case 'insertSheet':
        out.splice(m.at, 0, undefined)
        break
      case 'deleteSheet':
        out.splice(m.s, 1)
        break
      case 'moveSheet': {
        const [x] = out.splice(m.s, 1)
        out.splice(m.at, 0, x)
        break
      }
      case 'renameSheet':
        break
    }
  return out
}

/** The runs of columns where two patterns (of tables `ta` and `tb`) give a cell other looks. */
function differingRuns(pa: Pattern, ta: LookTable, pb: Pattern, tb: LookTable): number[] {
  const out: number[] = []
  for (const [c1, c2, a, b] of spansOf(pa, pb)) {
    if (sameLook(ta.list[a], tb.list[b])) continue
    const n = out.length
    if (n && out[n - 1] === c1 - 1) out[n - 1] = c2
    else out.push(c1, c2)
  }
  return out
}

/** The rectangles of cells to which two layers (none for no looks) give other looks: a run of columns over a block of rows, row by row. */
export function* differing(a: Layer | undefined, b: Layer | undefined): Generator<Rect> {
  if (a === b) return
  const cuts = new Set<number>()
  for (const x of [a, b]) for (let k = 0; k < (x?.blocks ?? 0); k++) cuts.add(x!.r1[k]).add(x!.r2[k] + 1)
  const rows = [...cuts].sort((x, y) => x - y)
  const empty = new LookTable()
  const memo = new Map<Pattern, Map<Pattern, number[]>>()
  for (let k = 0; k + 1 < rows.length; k++) {
    const [lo, hi] = [rows[k], rows[k + 1] - 1]
    const [pa, pb] = [a?.row(lo) ?? EMPTY, b?.row(lo) ?? EMPTY]
    if (pa === pb && a?.looks === b?.looks) continue
    let byB = memo.get(pa)
    if (!byB) memo.set(pa, (byB = new Map()))
    let runs = byB.get(pb)
    if (!runs) byB.set(pb, (runs = differingRuns(pa, a?.looks ?? empty, pb, b?.looks ?? empty)))
    for (let x = 0; x < runs.length; x += 2) yield { r1: lo, r2: hi, c1: runs[x], c2: runs[x + 1] }
  }
}
