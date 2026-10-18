/** Counters of the engine's work for tests of its cost: the formulas it evaluates and the cells of a column it reads. */
import type { EngineState } from '../src/engine.js'

/** Counts the formulas `state` evaluates from now on: `n`, set back to 0 by the test as it needs. */
export function evaluations(state: EngineState): { n: number } {
  const out = { n: 0 }
  for (const f of state.formulas) {
    const fn = f!.fn
    f!.fn = () => {
      out.n++
      return fn()
    }
  }
  return out
}

/** Counts the reads of the cells of column `col` of the first sheet of `state` from now on: the count so far. */
export function columnReads(state: EngineState, col: number): () => number {
  let reads = 0
  const column = state.ctx.sheets[0].cols[col]!
  const get = column.get.bind(column)
  column.get = (r) => {
    reads++
    return get(r)
  }
  return () => reads
}
