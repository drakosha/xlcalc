/**
 * A layer as the rectangles of `SheetInput.looks` and back. Any list is read, each rectangle over those before it. A
 * layer is written in one form for one set of looks (what a book gives back, and what its hash reads alike): blocks of
 * rows in order, a group of them over one base, the runs of the base's pattern as rectangles across the group's rows,
 * then each block of the group that differs from the base by the runs that differ, and the rows of the group without
 * looks as rows with none. A block joins the group when it differs from the base in less than half its runs, so rows
 * of one pattern apart from each other (rows in threes with empty rows between) take the base's runs once.
 */
import type { LookRange } from '../contract.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { areaText, parseRange, type Rect } from '../structure/area.js'
import { Layer, LayerBuilder, LookTable, NONE, patternOf, type Pattern } from './layer.js'

/** The runs of `q` over the columns where it differs from `base`: `[c1, c2, look]`, NONE where `q` has no look. */
export function difference(base: Pattern, q: Pattern): number[] {
  const out: number[] = []
  const lookAt = (p: Pattern, k: number) => (k < p.length ? p[k + 2] : NONE)
  // walk both patterns by the columns where either changes
  let [i, j, c] = [0, 0, 0]
  while (c < MAX_COLS) {
    while (i < base.length && base[i + 1] < c) i += 3
    while (j < q.length && q[j + 1] < c) j += 3
    const inBase = i < base.length && base[i] <= c
    const inQ = j < q.length && q[j] <= c
    const a = inBase ? lookAt(base, i) : NONE
    const b = inQ ? lookAt(q, j) : NONE
    // the next column where either pattern changes
    const end = Math.min(inBase ? base[i + 1] + 1 : i < base.length ? base[i] : MAX_COLS, inQ ? q[j + 1] + 1 : j < q.length ? q[j] : MAX_COLS)
    if (a !== b) {
      const n = out.length
      if (n && out[n - 1] === b && out[n - 2] === c - 1) out[n - 2] = end - 1
      else out.push(c, end - 1, b)
    }
    c = end
  }
  return out
}

/** The number of runs `top` has. */
const runs = (top: ArrayLike<number>) => top.length / 3

const rangeOf = (r1: number, r2: number, c1: number, c2: number) => areaText({ r1, r2, c1, c2 })

/** The rectangles of a layer, in the one form a layer is written in (see the module's comment). */
export function writeRanges(layer: Layer): LookRange[] {
  const out: LookRange[] = []
  const put = (r1: number, r2: number, c1: number, c2: number, look: number) => {
    const x = look === NONE ? undefined : layer.looks.list[look]
    out.push(x ? { range: rangeOf(r1, r2, c1, c2), ...x } : { range: rangeOf(r1, r2, c1, c2) })
  }
  const n = layer.blocks
  let b = 0
  while (b < n) {
    // a group: the base block, and the blocks after it that differ from it little, with the rows between them
    const base = layer.patterns[layer.at[b]]
    const members: [lo: number, hi: number, top: number[] | null][] = [[layer.r1[b], layer.r2[b], null]]
    let e = b + 1
    for (; e < n; e++) {
      const q = layer.patterns[layer.at[e]]
      const top = q === base ? null : difference(base, q)
      if (top && 2 * runs(top) >= runs(q)) break
      members.push([layer.r1[e], layer.r2[e], top])
    }
    const [g1, g2] = [layer.r1[b], layer.r2[e - 1]]
    for (let k = 0; k < base.length; k += 3) put(g1, g2, base[k], base[k + 1], base[k + 2])
    let next = g1
    for (const [lo, hi, top] of members) {
      if (lo > next) put(next, lo - 1, 0, MAX_COLS - 1, NONE)
      if (top) for (let k = 0; k < top.length; k += 3) put(lo, hi, top[k], top[k + 1], top[k + 2])
      next = hi + 1
    }
    b = e
  }
  return out
}

/** A rectangle of a list, parsed: its rows, columns, look and place in the list. */
interface Placed extends Rect {
  look: number
  order: number
}

/** A max-heap of rectangles by their place in the list, those taken out left in until they reach the top. */
class Tops {
  private readonly heap: Placed[] = []
  private readonly gone = new Set<Placed>()

  add(x: Placed) {
    const h = this.heap
    h.push(x)
    for (let i = h.length - 1; i > 0; ) {
      const p = (i - 1) >> 1
      if (h[p].order >= h[i].order) break
      ;[h[p], h[i]] = [h[i], h[p]]
      i = p
    }
  }

  remove(x: Placed) {
    this.gone.add(x)
  }

  /** The rectangle on top, the latest in the list. */
  top(): Placed | undefined {
    const h = this.heap
    while (h.length && this.gone.has(h[0])) {
      this.gone.delete(h[0])
      const last = h.pop()!
      if (!h.length) break
      h[0] = last
      for (let i = 0; ; ) {
        const [l, r] = [2 * i + 1, 2 * i + 2]
        let m = i
        if (l < h.length && h[l].order > h[m].order) m = l
        if (r < h.length && h[r].order > h[m].order) m = r
        if (m === i) break
        ;[h[m], h[i]] = [h[i], h[m]]
        i = m
      }
    }
    return h[0]
  }
}

/** The pattern rectangles `xs` (all over the row band at hand) paint, each over those before it. */
function paint(xs: Placed[]): Pattern {
  const events: [col: number, enter: boolean, x: Placed][] = []
  for (const x of xs) events.push([x.c1, true, x], [x.c2 + 1, false, x])
  events.sort((a, b) => a[0] - b[0])
  const tops = new Tops()
  const out: number[] = []
  for (let k = 0; k < events.length; ) {
    const c = events[k][0]
    for (; k < events.length && events[k][0] === c; k++) {
      const [, enter, x] = events[k]
      if (enter) tops.add(x)
      else tops.remove(x)
    }
    const top = tops.top()
    const end = k < events.length ? events[k][0] - 1 : MAX_COLS - 1
    if (top && top.look !== NONE && end >= c) out.push(c, end, top.look)
  }
  return patternOf(out)
}

/** The place in a sorted list where `v` is, or would go. */
function placeIn(list: number[], v: number): number {
  let [lo, hi] = [0, list.length]
  while (lo < hi) {
    const m = (lo + hi) >> 1
    if (list[m] < v) lo = m + 1
    else hi = m
  }
  return lo
}

/** The layer a list of rectangles (checked) gives, each over those before it, with its looks numbered by `looks`. */
export function readRanges(list: readonly LookRange[], looks = new LookTable()): Layer {
  const xs: Placed[] = list.map((x, order) => ({ ...parseRange(x.range)!, look: looks.id(x.style, x.numberFormat), order }))
  const out = new LayerBuilder(looks)
  // row bands between the rows where a rectangle starts or ends, each painted by the rectangles over it (by their place
  // in the list, in order); bands under the same rectangles share a pattern, and a band whose last rectangle spans every
  // column takes its look alone
  const edges = new Map<number, [enter: number[], leave: number[]]>()
  const edge = (r: number) => edges.get(r) ?? edges.set(r, [[], []]).get(r)!
  for (const x of xs) {
    edge(x.r1)[0].push(x.order)
    edge(x.r2 + 1)[1].push(x.order)
  }
  const rows = [...edges.keys()].sort((a, b) => a - b)
  const active: number[] = []
  const painted = new Map<string, Pattern>()
  for (let k = 0; k + 1 < rows.length; k++) {
    const [enter, leave] = edges.get(rows[k])!
    for (const o of leave) active.splice(placeIn(active, o), 1)
    for (const o of enter) active.splice(placeIn(active, o), 0, o)
    if (!active.length) continue
    const last = xs[active[active.length - 1]]
    let p: Pattern
    if (last.c1 === 0 && last.c2 === MAX_COLS - 1) p = patternOf([0, MAX_COLS - 1, last.look])
    else {
      const key = active.join(',')
      p = painted.get(key) ?? painted.set(key, paint(active.map((o) => xs[o]))).get(key)!
    }
    out.rows(rows[k], Math.min(rows[k + 1] - 1, MAX_ROWS - 1), p)
  }
  return out.build()
}
