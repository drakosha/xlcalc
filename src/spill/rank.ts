/**
 * Which of two arrays over one cell spills. Excel computes a formula after the formulas it reads, and the one
 * computed first spills; of two that do not depend on each other the winner varies between runs, the one lower
 * on the sheet winning most often (`docs/spill.md`). Here the order is the references' topological order with
 * ties taken by address from the end (sheet, row, column), so the formula another one reads wins, and otherwise
 * the one further down. Before that, an array whose formula reads, itself or through the formulas it reads, a cell
 * its array may spill over (one below and to the right of it on its sheet) gives way to one that reads none, as Excel
 * gives a workbook opened from its file where such an array closes a cycle through its spill (the workbook set
 * `spill-cycles-held`). Both tests follow the formulas and their references, not which of the two arrays spilled.
 */
import type { EngineState } from '../engine.js'
import { indegrees, releaseReady } from '../graph/order.js'

/** Formula `a` spills where its array and `b`'s meet. */
export function beats(state: EngineState, a: number, b: number): boolean {
  const lb = looped(state, b)
  if (looped(state, a) !== lb) return lb
  const rank = (state.spills.rank ??= ranks(state))
  // of two formulas in a cycle of references, neither ranked, the last by address
  return ((rank.get(a) ?? Infinity) - (rank.get(b) ?? Infinity) || before(state, a, b)) < 0
}

/**
 * Whether formula `id` reads, by its references (a spill reference as its cell and the rectangles around it), a cell below
 * and to the right of it on its sheet (its own included), itself or through the formulas it comes after in the graph: a
 * cycle through its spill wherever the array reaches that cell. Kept until the graph is built again.
 */
function looped(state: EngineState, id: number): boolean {
  const sp = state.spills
  const g = state.formulas[id]!
  const known = (sp.looped ??= new Map())
  if (!known.has(id)) {
    // the graph's edges turned round, made once for the graph
    const pred = (sp.pred ??= state.succ.reduce<number[][]>((p, next, v) => (next.forEach((n) => (p[n] ??= []).push(v)), p), []))
    const seen = new Set([id])
    let out = false
    // a set's walk takes in what is added to it on the way
    for (const v of seen) {
      const f = state.formulas[v]
      if (f) for (const x of [...f.refs, ...(f.spans ?? []).flatMap((y) => y.boxes)]) out ||= x.s === g.s && x.r2 >= g.r && x.c2 >= g.c
      if (out) break
      for (const n of pred[v] ?? []) seen.add(n)
    }
    known.set(id, out)
  }
  return known.get(id)!
}

/** Positions in the topological order of the references (without spills), the last by address among the ready ones. */
function ranks(state: EngineState): Map<number, number> {
  const { formulas, succ } = state
  const indeg = indegrees(succ)
  const heap = new Heap((a, b) => before(state, a, b))
  formulas.forEach((f, id) => f && indeg[id] === 0 && heap.push(id))
  const ready = (id: number) => heap.push(id)
  const out = new Map<number, number>()
  while (heap.size) {
    const id = heap.pop()
    out.set(id, out.size)
    releaseReady(id, succ, indeg, formulas.length, ready)
  }
  return out
}

/** Below 0 where formula `a` comes before `b` by address from the end (sheet, row, column). */
function before({ formulas }: EngineState, a: number, b: number) {
  const [x, y] = [formulas[a]!, formulas[b]!]
  return y.s - x.s || y.r - x.r || y.c - x.c
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
