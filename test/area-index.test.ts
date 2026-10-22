import { expect, it } from 'vitest'
import { AreaIndex } from '../src/area-index.js'
import type { Area } from '../src/spill/state.js'
import { meets } from '../src/structure/area.js'
import { mulberry32 } from '../src/synth.js'

type Entry = Area & { id: number }

/** The index as one set per column and one per sheet of wide entries: the order every lookup must keep. */
class Plain {
  columns = new Map<number, Map<number, Set<Entry>>>()
  wide = new Map<number, Set<Entry>>()
  add(e: Entry) {
    if (e.c2 - e.c1 >= 64) return void this.wide.set(e.s, (this.wide.get(e.s) ?? new Set()).add(e))
    const cols = this.columns.get(e.s) ?? new Map<number, Set<Entry>>()
    this.columns.set(e.s, cols)
    for (let c = e.c1; c <= e.c2; c++) cols.set(c, (cols.get(c) ?? new Set()).add(e))
  }
  delete(e: Entry) {
    this.wide.get(e.s)?.delete(e)
    for (let c = e.c1; c <= e.c2; c++) {
      const set = this.columns.get(e.s)?.get(c)
      set?.delete(e)
      if (set && !set.size) this.columns.get(e.s)!.delete(c)
    }
  }
  forEach(a: Area, visit: (e: Entry) => void) {
    const scan = (set?: Set<Entry>) => set && [...set].forEach((e) => meets(e, a) && visit(e))
    const cols = this.columns.get(a.s)
    if (cols && a.c2 - a.c1 + 1 > cols.size) for (const [c, set] of cols) c >= a.c1 && c <= a.c2 && scan(set)
    else if (cols) for (let c = a.c1; c <= a.c2; c++) scan(cols.get(c))
    scan(this.wide.get(a.s))
  }
}

it('finds what meets an area in the order a set per column lists it, through additions and deletions', () => {
  const next = mulberry32(7), rand = (n: number) => Math.floor(next() * n)
  const span = (max: number, long: number) => {
    const lo = rand(max), size = rand(4) ? rand(long) : rand(max)
    return [lo, Math.min(max - 1, lo + size)]
  }
  const area = (): Area => {
    const [r1, r2] = span(600, rand(2) ? 3 : 140), [c1, c2] = span(90, rand(3) ? 2 : 80)
    return { s: rand(2), r1, r2, c1, c2 }
  }
  const index = new AreaIndex<Entry>(), plain = new Plain(), live: Entry[] = []
  for (let step = 0; step < 6000; step++) {
    if (live.length && rand(4) === 0) {
      const [e] = live.splice(rand(live.length), 1)
      index.delete(e)
      plain.delete(e)
    } else {
      const e = { ...area(), id: step }
      live.push(e)
      index.add(e)
      plain.add(e)
    }
    const a = area(), got: number[] = [], want: number[] = []
    index.forEach(a, (e) => got.push(e.id))
    plain.forEach(a, (e) => want.push(e.id))
    expect(got).toEqual(want)
  }
})

it('lets a visitor delete what it visits', () => {
  const index = new AreaIndex<Entry>()
  const entries: Entry[] = [{ s: 0, r1: 60, r2: 70, c1: 0, c2: 1, id: 0 }, { s: 0, r1: 0, r2: 500, c1: 1, c2: 1, id: 1 }, { s: 0, r1: 65, r2: 65, c1: 1, c2: 1, id: 2 }]
  for (const e of entries) index.add(e)
  const seen: number[] = []
  index.forEach({ s: 0, r1: 65, r2: 65, c1: 0, c2: 1 }, (e) => { seen.push(e.id); index.delete(e) })
  expect(seen).toEqual([0, 1, 2])
  index.forEach({ s: 0, r1: 0, r2: 1000, c1: 0, c2: 5 }, (e) => seen.push(e.id))
  expect(seen).toEqual([0, 1, 2])
})
