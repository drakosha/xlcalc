/**
 * Which of two arrays over one cell spills. Excel computes a formula after the formulas it reads, and the one
 * computed first spills; of two that do not depend on each other the winner varies between runs, the one lower
 * on the sheet winning most often (`docs/spill.md`). Here the order is the references' topological order with
 * ties taken by address from the end (sheet, row, column), so the formula another one reads wins, and otherwise
 * the one further down.
 */
import type { EngineState } from '../engine.js'

/** Formula `a` spills where its array and `b`'s meet. */
export function beats(state: EngineState, a: number, b: number): boolean {
  const rank = (state.spills.rank ??= ranks(state))
  return (rank.get(a) ?? Infinity) < (rank.get(b) ?? Infinity)
}

/** Positions in the topological order of the references (without spills), the last by address among the ready ones. */
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
  const indeg = new Array<number>(formulas.length).fill(0)
  for (const f of formulas) if (f) for (const n of succ[f.id]) indeg[n]++
  const heap = new Heap(before)
  formulas.forEach((f, id) => f && indeg[id] === 0 && heap.push(id))
  const out = new Map<number, number>()
  while (heap.size) {
    const id = heap.pop()
    out.set(id, out.size)
    for (const n of succ[id]) if (--indeg[n] === 0) heap.push(n)
  }
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
