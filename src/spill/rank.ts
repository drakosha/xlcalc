/**
 * Which of two arrays over one cell spills. Excel computes a formula after the formulas it reads, and the one
 * computed first spills; of two that do not depend on each other the winner varies between runs, the one lower
 * on the sheet winning most often (`docs/spill.md`). Here the order is the references' topological order with
 * ties taken by address from the end (sheet, row, column), so the formula another one reads wins, and otherwise
 * the one further down. An array whose range its own formula reads back, a cycle through its spill, gives way to
 * one that closes no such cycle, as Excel gives a workbook opened from its file (the workbook set `spill-cycles-held`):
 * the outcome depends on neither the order of the formulas nor which one was computed first.
 */
import type { EngineState } from '../engine.js'
import { readersOf } from '../graph.js'
import { indegrees, releaseReady } from '../graph/order.js'
import type { Area } from './state.js'

/** Formula `a` spills where its array (over `area`, else the range it holds or waits for) and `b`'s meet. */
export function beats(state: EngineState, a: number, b: number, area?: Area): boolean {
  const [la, lb] = [looped(state, a, area), looped(state, b)]
  if (la !== lb) return lb
  const rank = (state.spills.rank ??= ranks(state))
  return rank.get(a)! < rank.get(b)!
}

/** The range formula `id` spilled into or keeps in a cycle (its claim), else the one it waits for. */
const areaOf = (state: EngineState, id: number): Area | undefined => state.spills.claims.get(id) ?? state.spills.wait.get(id)

/**
 * Whether formula `id`'s array over `area` (else its own range) is read back by the formula: a path of references from a
 * reader of the range to it. The ranges of other arrays are left out, as they change with what spilled.
 */
function looped(state: EngineState, id: number, area = areaOf(state, id)): boolean {
  if (!area) return false
  const seen = new Set<number>()
  readersOf(state, area, seen)
  const stack = [...seen]
  while (stack.length) {
    const v = stack.pop()!
    if (v === id) return true
    for (const n of state.succ[v] ?? []) if (!seen.has(n)) {
      seen.add(n)
      stack.push(n)
    }
  }
  return false
}

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
