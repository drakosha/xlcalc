/**
 * Which of two arrays over one cell spills. Excel computes a formula after the formulas it reads, and the one
 * computed first spills; of two that do not depend on each other the winner varies between runs, the one lower
 * on the sheet winning most often (`docs/spill.md`). Here the order is the references' topological order with
 * ties taken by address from the end (sheet, row, column), so the formula another one reads wins, and otherwise
 * the one further down. Before that, an array whose formula reads, itself or through the formulas it reads, a cell
 * its array may spill over (one below and to the right of it on its sheet) gives way to one that reads none, as Excel
 * gives a workbook opened from its file where such an array closes a cycle through its spill (the workbook set
 * `spill-cycles-held`). Both tests follow the formulas and their references alone, not what has spilled.
 */
import type { EngineState, Formula } from '../engine.js'
import { spend } from '../exec/budget.js'
import { indegrees, releaseReady } from '../graph/order.js'
import { lowerBound } from '../sorted.js'
import type { Rng } from '../values.js'

/** Formula `a` spills where its array and `b`'s meet. */
export function beats(state: EngineState, a: number, b: number): boolean {
  const [la, lb] = [looped(state, a), looped(state, b)]
  if (la !== lb) return lb
  const rank = (state.spills.rank ??= ranks(state))
  return (rank.get(a) ?? Infinity) < (rank.get(b) ?? Infinity)
}

/** The ranges a formula reads by its references, a spill reference as its cell and the rectangles around it. */
const reads = (f: Formula): readonly Rng[] => (f.spans ? [...f.refs, ...f.spans.flatMap((x) => x.boxes)] : f.refs)

/**
 * Whether formula `id` reads a cell below and to the right of it on its sheet (its own included), itself or through
 * the formulas in the ranges it reads: a cycle through its spill wherever the array reaches that cell. Kept until the graph
 * is built again.
 */
function looped(state: EngineState, id: number): boolean {
  const sp = state.spills
  const known = (sp.looped ??= new Map())
  let out = known.get(id)
  if (out !== undefined) return out
  const g = state.formulas[id]!
  const places = (sp.places ??= placesOf(state))
  const seen = new Set([id])
  const stack = [id]
  out = false
  while (stack.length && !out) {
    for (const x of reads(state.formulas[stack.pop()!]!)) {
      if (x.s === g.s && x.r2 >= g.r && x.c2 >= g.c) out = true
      else places.each(x, (n) => !seen.has(n) && (seen.add(n), stack.push(n)))
    }
  }
  known.set(id, out)
  return out
}

/** The formulas by sheet, column and row, to find those in a range. */
export class Places {
  private readonly sheets = new Map<number, { cols: number[]; rows: Map<number, { r: number; id: number }[]> }>()

  constructor(formulas: readonly (Formula | undefined)[]) {
    for (const f of formulas) {
      if (!f) continue
      let sh = this.sheets.get(f.s)
      if (!sh) this.sheets.set(f.s, (sh = { cols: [], rows: new Map() }))
      let rows = sh.rows.get(f.c)
      if (!rows) {
        sh.rows.set(f.c, (rows = []))
        sh.cols.push(f.c)
      }
      rows.push({ r: f.r, id: f.id })
    }
    for (const sh of this.sheets.values()) {
      sh.cols.sort((a, b) => a - b)
      for (const rows of sh.rows.values()) rows.sort((a, b) => a.r - b.r)
    }
  }

  /** Calls `fn` with each formula in the range. */
  each(x: Rng, fn: (id: number) => unknown) {
    const sh = this.sheets.get(x.s)
    if (!sh) return
    for (let i = lowerBound(sh.cols, x.c1, (c) => c); i < sh.cols.length && sh.cols[i] <= x.c2; i++) {
      const rows = sh.rows.get(sh.cols[i])!
      for (let j = lowerBound(rows, x.r1, (y) => y.r); j < rows.length && rows[j].r <= x.r2; j++) {
        spend(1)
        fn(rows[j].id)
      }
    }
  }
}

const placesOf = (state: EngineState) => new Places(state.formulas)

/**
 * Positions in the topological order of the references (without spills), the last by address among the ready ones;
 * the formulas in a cycle of references or over one after all the others, the last by address first.
 */
function ranks(state: EngineState): Map<number, number> {
  const { formulas, succ } = state
  const place = (id: number) => {
    const f = formulas[id]!
    return [f.s, f.r, f.c]
  }
  const before = (a: number, b: number) => {
    const [x, y] = [place(a), place(b)]
    return y[0] - x[0] || y[1] - x[1] || y[2] - x[2]
  }
  const indeg = indegrees(succ)
  const heap = new Heap(before)
  formulas.forEach((f, id) => f && indeg[id] === 0 && heap.push(id))
  const ready = (id: number) => heap.push(id)
  const out = new Map<number, number>()
  while (heap.size) {
    const id = heap.pop()
    out.set(id, out.size)
    releaseReady(id, succ, indeg, formulas.length, ready)
  }
  const rest: number[] = []
  formulas.forEach((f, id) => f && !out.has(id) && rest.push(id))
  for (const id of rest.sort(before)) out.set(id, out.size)
  return out
}

/** A binary min-heap of ids. */
class Heap {
  private readonly xs: number[] = []
  constructor(private readonly less: (a: number, b: number) => number) {}

  get size() {
    return this.xs.length
  }

  push(x: number) {
    const xs = this.xs
    xs.push(x)
    for (let i = xs.length - 1; i > 0; ) {
      const p = (i - 1) >> 1
      if (this.less(xs[i], xs[p]) >= 0) break
      ;[xs[i], xs[p]] = [xs[p], xs[i]]
      i = p
    }
  }

  pop(): number {
    const xs = this.xs
    const top = xs[0]
    const last = xs.pop()!
    if (xs.length) {
      xs[0] = last
      for (let i = 0; ; ) {
        const [l, r] = [2 * i + 1, 2 * i + 2]
        let m = i
        if (l < xs.length && this.less(xs[l], xs[m]) < 0) m = l
        if (r < xs.length && this.less(xs[r], xs[m]) < 0) m = r
        if (m === i) break
        ;[xs[i], xs[m]] = [xs[m], xs[i]]
        i = m
      }
    }
    return top
  }
}
