/** The values a structure operation changed in a session, and the formulas and looks: compared with those of the cell that moved there. */
import type { CellInput } from '../contract.js'
import { key, type EngineState } from '../engine.js'
import { same } from '../evaluate.js'
import { spend } from '../exec/budget.js'
import { formatAddress, MAX_COLS } from '../model.js'
import { fieldsBetween, formulaText, type FieldChange } from './answer.js'
import type { Looks } from '../shown.js'
import { shownKey } from './looks.js'
import { describeAt } from '../spill/place.js'
import { inverse, placer, type Geometry, type Place } from '../structure/geometry.js'

/**
 * Cells of the new state whose value differs from the one their old cell had (the geometry `g` carries each old value
 * to its new place; a cell nothing came to was empty), and formulas whose spill differs from their old cell's
 * (`describeAt`, which stands for a large array's unstored cells), as `s:r:c` keys of the new state in order; with the
 * looks of both states, also the cells that show something else than their old cell showed. With `fields`, also the
 * cells whose formula (as each state shows it), own style or own number format differ from their old cell's, by key,
 * found in the same pass. Each cell is found by its numbers both ways: where its old cell went, and where a new cell
 * came from (the geometry's inverse), so no cell is looked up by its key.
 */
export function movedChanges(old: EngineState, next: EngineState, g: Geometry, looks?: [Looks, Looks], fields?: Map<string, FieldChange>): string[] {
  const [place, back] = [placer(g), placer(inverse(g))]
  const p: Place = { s: 0, r: 0, c: 0 }
  const input = (st: EngineState, s: number, r: number, c: number) => st.input.sheets[s].cells[formatAddress(r, c)]
  const textOf = (st: EngineState, s: number, r: number, c: number, x: CellInput | undefined) => (x?.formula === undefined ? undefined : formulaText(st, s, r, c, x))
  /** The fields of the new cell at `j`, `r`, `c` against those of the old cell at `from` that came to it (none: an empty one). */
  const compare = (j: number, r: number, c: number, from?: Place) => {
    const now = input(next, j, r, c)
    const was = from && input(old, from.s, from.r, from.c)
    // the same input without a formula: nothing of it can differ (a formula's text may, with new names)
    if (was === now && was?.formula === undefined) return
    const f = fieldsBetween(was, from && textOf(old, from.s, from.r, from.c, was), now, textOf(next, j, r, c, now))
    if (f) fields!.set(key(j, r, c), f)
  }
  const valueAt = (st: EngineState, s: number, r: number, c: number) => st.ctx.sheets[s].cols[c]?.[r]
  const bare = (st: EngineState, s: number, r: number, c: number) => st.bare.get(s)?.has(r * MAX_COLS + c) === true
  const out: number[][] = []
  const from: Place = { s: 0, r: 0, c: 0 }
  const at = (x: Place, s: number, r: number, c: number) => {
    x.s = s
    x.r = r
    x.c = c
    return x
  }
  // each old value where the geometry takes it, against the new value there
  old.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        if (v === null) return
        at(p, s, r, c)
        if (!place(p)) return
        const w = valueAt(next, p.s, p.r, p.c) ?? null
        if (fields) compare(p.s, p.r, p.c, at(from, s, r, c))
        if (!same(v, w) || (looks && w !== null && shownKey(looks[0], old, s, r, c) !== shownKey(looks[1], next, p.s, p.r, p.c))) out.push([p.s, p.r, p.c])
      }),
    ),
  )
  // the cells of a style, a format, a note or a link alone, which hold no value
  if (fields)
    for (const [s, cells] of old.bare)
      for (const n of cells) {
        at(p, s, Math.floor(n / MAX_COLS), n % MAX_COLS)
        at(from, p.s, p.r, p.c)
        if (place(p)) compare(p.s, p.r, p.c, from)
      }
  /** Whether the old cell at `p` (where `back` took a new one) holds a value. */
  const valued = () => (valueAt(old, p.s, p.r, p.c) ?? null) !== null
  // the new values no old value came to, against an empty cell
  next.ctx.sheets.forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        const came = back(at(p, s, r, c))
        if (came && valued()) return
        // an old cell of a look alone that came to it was compared above
        if (fields && !(came && bare(old, p.s, p.r, p.c))) compare(s, r, c)
        if (v !== null) out.push([s, r, c])
      }),
    ),
  )
  if (fields)
    for (const [s, cells] of next.bare)
      for (const n of cells) {
        const [r, c] = [Math.floor(n / MAX_COLS), n % MAX_COLS]
        // compared above: an old cell of a look alone came to it, or one with a value where the new state has none
        if (back(at(p, s, r, c)) && (bare(old, p.s, p.r, p.c) || (valued() && valueAt(next, s, r, c) === undefined))) continue
        compare(s, r, c)
      }
  const spills = new Map<string, string>()
  for (const x of old.spills.on.values()) {
    at(p, x.s, x.r1, x.c1)
    if (place(p)) spills.set(key(p.s, p.r, p.c), describeAt(old, key(x.s, x.r1, x.c1)))
  }
  for (const x of next.spills.on.values()) spills.set(key(x.s, x.r1, x.c1), spills.get(key(x.s, x.r1, x.c1)) ?? '')
  let listed: Set<string> | undefined
  for (const [k, was] of spills)
    if (was !== describeAt(next, k) && !(listed ??= new Set(out.map((o) => o.join(':')))).has(k)) out.push(k.split(':').map(Number))
  return out.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]).map((k) => k.join(':'))
}
