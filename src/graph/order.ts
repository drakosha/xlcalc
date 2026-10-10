/** Topological bookkeeping shared by evaluation order and deterministic spill priority. */
import { spend } from '../exec/budget.js'

export function indegrees(succ: readonly number[][], extra?: ReadonlyMap<number, number[]>): number[] {
  const degrees = new Array<number>(succ.length).fill(0)
  for (const next of succ) {
    spend(next.length)
    for (const id of next) degrees[id]++
  }
  if (extra) for (const next of extra.values()) for (const id of next) degrees[id]++
  return degrees
}

/** Finish one formula and drain all newly ready helper vertices before choosing another formula. */
export function releaseReady(id: number, succ: readonly number[][], degrees: number[], formulas: number, ready: (id: number) => void, extra?: ReadonlyMap<number, number[]>): void {
  const stack = [id]
  const edge = (next: number) => {
    if (--degrees[next] === 0) {
      if (next < formulas) ready(next)
      else stack.push(next)
    }
  }
  while (stack.length) {
    const current = stack.pop()!
    spend(succ[current].length)
    for (const next of succ[current]) edge(next)
    for (const next of extra?.get(current) ?? []) edge(next)
  }
}
