/** The reference graph over a book's formulas: which formulas read which cells, and the order of evaluation. */
import type { EngineState, Formula } from './engine.js'

const WIDE = 64

/** The reference graph of the formulas: successors, the topological order, listeners of ranges, volatile formulas. */
export function rebuildGraph(state: EngineState) {
  const { formulas } = state
  const at = new Map<number, Map<number, { r: number; id: number }[]>>()
  for (const f of formulas) {
    if (!f) continue
    let bySheet = at.get(f.s)
    if (!bySheet) at.set(f.s, (bySheet = new Map()))
    let col = bySheet.get(f.c)
    if (!col) bySheet.set(f.c, (col = []))
    col.push({ r: f.r, id: f.id })
  }
  for (const bySheet of at.values()) for (const col of bySheet.values()) col.sort((a, b) => a.r - b.r)
  const lowerBound = (col: { r: number }[], r: number) => {
    let lo = 0
    let hi = col.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (col[mid].r < r) lo = mid + 1
      else hi = mid
    }
    return lo
  }
  const succ: number[][] = formulas.map(() => [])
  const indeg = new Array<number>(formulas.length).fill(0)
  const listeners: EngineState['listeners'] = new Map()
  const wide: EngineState['wide'] = new Map()
  for (const f of formulas) {
    if (!f) continue
    for (const g of f.refs) {
      if (g.c2 - g.c1 >= WIDE) {
        let w = wide.get(g.s)
        if (!w) wide.set(g.s, (w = []))
        w.push([g.r1, g.r2, g.c1, g.c2, f.id])
      } else {
        let ls = listeners.get(g.s)
        if (!ls) listeners.set(g.s, (ls = new Map()))
        for (let c = g.c1; c <= g.c2; c++) {
          let lc = ls.get(c)
          if (!lc) ls.set(c, (lc = []))
          lc.push([g.r1, g.r2, f.id])
        }
      }
      const bySheet = at.get(g.s)
      if (!bySheet) continue
      for (const [c, col] of bySheet) {
        if (c < g.c1 || c > g.c2) continue
        for (let i = lowerBound(col, g.r1); i < col.length && col[i].r <= g.r2; i++) {
          succ[col[i].id].push(f.id)
          indeg[f.id]++
        }
      }
    }
  }
  const order: number[] = []
  const queue: number[] = []
  formulas.forEach((f, id) => f && indeg[id] === 0 && queue.push(id))
  // the queue need not be ordered by address: values do not depend on the evaluation order
  while (queue.length) {
    const id = queue.pop()!
    order.push(id)
    for (const n of succ[id]) if (--indeg[n] === 0) queue.push(n)
  }
  const topo = new Array<number>(formulas.length).fill(-1)
  order.forEach((id, i) => (topo[id] = i))
  state.succ = succ
  state.topo = topo
  state.order = order
  state.listeners = listeners
  state.wide = wide
  const vol = new Set<number>()
  const stack = formulas.filter((f): f is Formula => !!f?.volatile).map((f) => f.id)
  while (stack.length) {
    const id = stack.pop()!
    if (vol.has(id)) continue
    vol.add(id)
    stack.push(...succ[id])
  }
  state.volatile = vol
}

/** Formulas that read a cell (by their references, whether or not the cell holds anything). */
export function listenersOf(state: EngineState, s: number, r: number, c: number, out: Set<number>) {
  for (const [r1, r2, id] of state.listeners.get(s)?.get(c) ?? []) if (r >= r1 && r <= r2) out.add(id)
  for (const [r1, r2, c1, c2, id] of state.wide.get(s) ?? []) if (r >= r1 && r <= r2 && c >= c1 && c <= c2) out.add(id)
}
