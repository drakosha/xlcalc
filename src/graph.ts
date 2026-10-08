/**
 * The reference graph over a book's formulas: which formulas read which cells, and the order of evaluation.
 * Besides the references, a spilled array (its claim, `spill/state.ts`) comes before the formulas reading its range.
 */
import type { EngineState, Formula } from './engine.js'
import { spend } from './exec/budget.js'
import { bySheet } from './exec/walk.js'
import type { Area } from './spill/state.js'
import { Rng } from './values.js'

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
  // the formula at a cell, if any
  const idAt = (s: number, r: number, c: number) => {
    const col = at.get(s)?.get(c)
    const i = col ? lowerBound(col, r) : -1
    return col && col[i]?.r === r ? col[i].id : undefined
  }
  // a range operator reaching a spilled array reads as far as the array claims, around its other ends
  const spanned = (f: Formula): readonly Rng[] => {
    if (!f.spans) return f.refs
    const more: Rng[] = []
    for (const span of f.spans) {
      const reached = [...span.boxes]
      for (const a of span.anchors) {
        const id = idAt(a.s, a.r1, a.c1)
        const x = id === undefined ? undefined : state.spills.claims.get(id)
        if (x) reached.push(new Rng(x.s, x.r1, x.c1, x.r2, x.c2))
      }
      if (reached.length > span.boxes.length) more.push(...bySheet(reached).values())
    }
    return more.length ? [...f.refs, ...more] : f.refs
  }
  const succ: number[][] = formulas.map(() => [])
  const indeg = new Array<number>(formulas.length).fill(0)
  const listeners: EngineState['listeners'] = new Map()
  const wide: EngineState['wide'] = new Map()
  for (const f of formulas) {
    if (!f) continue
    spend(1)
    for (const g of spanned(f)) {
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
  state.succ = succ
  state.listeners = listeners
  state.wide = wide
  const spillSucc = new Map<number, number[]>()
  for (const [id, x] of state.spills.claims) {
    const readers = new Set<number>()
    readersOf(state, x, readers)
    if (!readers.size) continue
    spillSucc.set(id, [...readers])
    for (const n of readers) indeg[n]++
  }
  state.spillSucc = spillSucc
  state.spills.rank = undefined
  const order: number[] = []
  const queue: number[] = []
  formulas.forEach((f, id) => f && indeg[id] === 0 && queue.push(id))
  // the queue need not be ordered by address: values do not depend on the evaluation order (of two arrays
  // over one cell, the one that spills is chosen by rank, `spill/rank.ts`)
  while (queue.length) {
    const id = queue.pop()!
    order.push(id)
    for (const n of successors(state, id)) if (--indeg[n] === 0) queue.push(n)
  }
  const topo = new Array<number>(formulas.length).fill(-1)
  order.forEach((id, i) => (topo[id] = i))
  state.topo = topo
  state.order = order
  markVolatile(state)
}

/** The formulas that depend on the clock: volatile functions and the year of dates without one (`clocked`), and the formulas over them. */
export function markVolatile(state: EngineState) {
  const vol = new Set<number>()
  const stack = [...state.formulas.filter((f): f is Formula => !!f?.volatile).map((f) => f.id), ...state.clocked]
  while (stack.length) {
    const id = stack.pop()!
    if (vol.has(id)) continue
    vol.add(id)
    stack.push(...successors(state, id))
  }
  state.volatile = vol
}

/** The formulas that come after a formula: those reading its cell, and those reading the range its array spilled into. */
export function successors(state: EngineState, id: number): number[] {
  const spilled = state.spillSucc.get(id)
  return spilled ? [...state.succ[id], ...spilled] : (state.succ[id] ?? [])
}

/** Formulas that read a cell of an area (by their references, whether or not the cells hold anything). */
export function readersOf(state: EngineState, x: Area, out: Set<number>) {
  const ls = state.listeners.get(x.s)
  const scan = (list: [number, number, number][]) => {
    for (const [r1, r2, id] of list) if (r1 <= x.r2 && r2 >= x.r1) out.add(id)
  }
  // a wide area is checked against the listened columns, a narrow one column by column
  if (ls && x.c2 - x.c1 + 1 > ls.size) for (const [c, list] of ls) c >= x.c1 && c <= x.c2 && scan(list)
  else if (ls) for (let c = x.c1; c <= x.c2; c++) scan(ls.get(c) ?? [])
  for (const [r1, r2, c1, c2, id] of state.wide.get(x.s) ?? []) if (r1 <= x.r2 && r2 >= x.r1 && c1 <= x.c2 && c2 >= x.c1) out.add(id)
}
