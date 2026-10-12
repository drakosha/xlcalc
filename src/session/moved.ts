/** The values a structure operation changed in a session, and the formulas and looks: compared with those of the cell that moved there. */
import type { CellInput } from '../contract.js'
import { key, type EngineState } from '../engine.js'
import { same } from '../evaluate.js'
import { spend } from '../exec/budget.js'
import { formatAddress, MAX_COLS, MAX_ROWS, readAddress } from '../model.js'
import { fieldsBetween, formulaText, hasFields, type FieldChange } from './answer.js'
import type { Looks } from '../shown.js'
import { shownKey } from './looks.js'
import type { DiffPass } from './rebook.js'
import { describeAt } from '../spill/place.js'
import { inverse, newSheet, placer, type Geometry, type Place } from '../structure/geometry.js'

/** A cell's place as one number, which orders places as addresses: by sheet, row, then column. */
const placeOf = (s: number, r: number, c: number) => (s * MAX_ROWS + r) * MAX_COLS + c

/** Places (`placeOf`) as `s:r:c` keys in address order, once each. */
function keysOf(places: number[]): string[] {
  places.sort((a, b) => a - b)
  const out: string[] = []
  let last = -1
  for (const n of places) {
    if (n === last) continue
    last = n
    const c = n % MAX_COLS
    const rest = (n - c) / MAX_COLS
    const r = rest % MAX_ROWS
    out.push(`${(rest - r) / MAX_ROWS}:${r}:${c}`)
  }
  return out
}

/**
 * Cells of the new state whose value differs from the one their old cell had (the geometry `g` carries each old value
 * to its new place; a cell nothing came to was empty), and formulas whose spill differs from their old cell's
 * (`describeAt`, which stands for a large array's unstored cells); with the looks of both states, also the cells that
 * show something else than their old cell showed. Into `fields`, the cells whose formula (as each state shows it), own
 * style or own number format differ from their old cell's, by key, found in the same pass. All of them as `s:r:c` keys
 * of the new state in address order, once each. Each cell is found by its numbers both ways: where its old cell went,
 * and where a new cell came from (the geometry's inverse), so no cell is looked up by its key; the cells of a sheet no
 * old cell comes to (a deleted sheet an undo puts back), by the addresses of its inputs.
 *
 * A cell is here one with a value (a spilled one too) or of a look alone (`EngineState.bare`); the book's diff
 * (`rebook.ts`, with `diff`) counts inputs, which a spilled cell is not. The pass gives the diff each pair of inputs it
 * looks up, once: an old cell of a value or a look alone with the new input at its place (or none, the geometry taking
 * it nowhere), and a new input that came from a place of neither, with the old input there. That meets every input, as
 * an input holds a value or is of a look alone.
 */
export function movedChanges(old: EngineState, next: EngineState, g: Geometry, fields: Map<string, FieldChange>, looks?: [Looks, Looks], diff?: DiffPass): string[] {
  const [place, back] = [placer(g), placer(inverse(g))]
  const p: Place = { s: 0, r: 0, c: 0 }
  const input = (st: EngineState, s: number, r: number, c: number) => st.input.sheets[s].cells[formatAddress(r, c)]
  const textOf = (st: EngineState, s: number, r: number, c: number, x: CellInput | undefined) => (x?.formula === undefined ? undefined : formulaText(st, s, r, c, x))
  /**
   * The fields of the new cell at `j`, `r`, `c` against those of the old cell at `from` that came to it (none: an empty
   * one); with `pair`, the two inputs go to the diff.
   */
  const compare = (j: number, r: number, c: number, from?: Place, pair = false) => {
    const now = input(next, j, r, c)
    const addr = from && formatAddress(from.r, from.c)
    const was = addr === undefined ? undefined : old.input.sheets[from!.s].cells[addr]
    if (pair) diff?.pair(from!.s, addr!, was, j, now)
    // the same input without a formula: nothing of it can differ (a formula's text may, with new names)
    if (was === now && was?.formula === undefined) return
    const f = fieldsBetween(was, from && textOf(old, from.s, from.r, from.c, was), now, textOf(next, j, r, c, now))
    if (!f) return
    fields.set(key(j, r, c), f)
    out.push(placeOf(j, r, c))
  }
  const valueAt = (st: EngineState, s: number, r: number, c: number) => st.ctx.sheets[s].cols[c]?.[r]
  const bare = (st: EngineState, s: number, r: number, c: number) => st.bare.get(s)?.has(r * MAX_COLS + c) === true
  /** Whether the old cell at `p` (where `back` took a new one) holds a value. */
  const valued = () => (valueAt(old, p.s, p.r, p.c) ?? null) !== null
  /** The new input at `s`, `r`, `c` and the old one at `p` it came from, a place of neither a value nor a look alone, to the diff. */
  const unmet = (s: number, r: number, c: number) => {
    if (!diff?.takes(p.s)) return
    const addr = formatAddress(p.r, p.c)
    diff.pair(p.s, addr, old.input.sheets[p.s].cells[addr], s, input(next, s, r, c))
  }
  // the places of the cells listed, sorted at the end
  const out: number[] = []
  const from: Place = { s: 0, r: 0, c: 0 }
  const at = (x: Place, s: number, r: number, c: number) => {
    x.s = s
    x.r = r
    x.c = c
    return x
  }
  // each old value where the geometry takes it, against the new value there; the linked workbooks' sheets after the
  // book's own (`links.ts`) move with nothing
  old.ctx.sheets.slice(0, old.sheetNames.length).forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        if (v === null) return
        if (!place(at(p, s, r, c))) {
          diff?.gone(s, r, c)
          return
        }
        const w = valueAt(next, p.s, p.r, p.c) ?? null
        compare(p.s, p.r, p.c, at(from, s, r, c), true)
        if (!same(v, w) || (looks && w !== null && shownKey(looks[0], old, s, r, c) !== shownKey(looks[1], next, p.s, p.r, p.c))) out.push(placeOf(p.s, p.r, p.c))
      }),
    ),
  )
  // the cells of a style, a format, a note or a link alone, which hold no value (one a spill covers was met above)
  for (const [s, cells] of old.bare)
    for (const n of cells) {
      const [r, c] = [Math.floor(n / MAX_COLS), n % MAX_COLS]
      const own = (valueAt(old, s, r, c) ?? null) === null
      if (place(at(p, s, r, c))) compare(p.s, p.r, p.c, at(from, s, r, c), own)
      else if (own) diff?.gone(s, r, c)
    }
  // a new sheet (one an undo puts back): no old cell comes to it, so its cells are compared with none by their inputs below
  const fresh = next.sheetNames.map((_, s) => newSheet(g, s))
  // the new values no old value came to, against an empty cell
  next.ctx.sheets.slice(0, next.sheetNames.length).forEach((sh, s) =>
    sh.cols.forEach((col, c) =>
      col?.forEach((v, r) => {
        spend(1)
        if (fresh[s]) {
          if (v !== null) out.push(placeOf(s, r, c))
          return
        }
        const came = back(at(p, s, r, c))
        if (came && valued()) return
        // an old cell of a look alone that came to it was compared above
        const alone = came && bare(old, p.s, p.r, p.c)
        // a new input from a place of no input (one of a look alone goes to the diff below)
        if (came && !alone && !bare(next, s, r, c)) unmet(s, r, c)
        if (!alone) compare(s, r, c)
        if (v !== null) out.push(placeOf(s, r, c))
      }),
    ),
  )
  for (const [s, cells] of next.bare) {
    if (fresh[s]) continue
    for (const n of cells) {
      const [r, c] = [Math.floor(n / MAX_COLS), n % MAX_COLS]
      const came = back(at(p, s, r, c))
      // a new input of a look alone, a spilled one too, from a place of no input
      if (came && !bare(old, p.s, p.r, p.c) && !valued()) unmet(s, r, c)
      // compared above: an old cell of a look alone came to it, or one with a value where the new state has none
      if (came && (bare(old, p.s, p.r, p.c) || (valued() && valueAt(next, s, r, c) === undefined))) continue
      compare(s, r, c)
    }
  }
  // the cells of a new sheet met above (with a value or of a look alone) whose input has a field beside its value
  // (`hasFields`): found by the addresses of its inputs, not each cell looked up
  fresh.forEach((yes, s) => {
    if (!yes) return
    const [cells, cols] = [next.input.sheets[s].cells, next.ctx.sheets[s].cols]
    for (const addr in cells) {
      if (!hasFields(cells[addr])) continue
      readAddress(addr, p)
      if ((cols[p.c] !== undefined && p.r in cols[p.c]) || bare(next, s, p.r, p.c)) compare(s, p.r, p.c)
    }
  })
  const spills = new Map<string, string>()
  for (const x of old.spills.on.values()) {
    at(p, x.s, x.r1, x.c1)
    if (place(p)) spills.set(key(p.s, p.r, p.c), describeAt(old, key(x.s, x.r1, x.c1)))
  }
  for (const x of next.spills.on.values()) spills.set(key(x.s, x.r1, x.c1), spills.get(key(x.s, x.r1, x.c1)) ?? '')
  for (const [k, was] of spills)
    if (was !== describeAt(next, k)) {
      const [s, r, c] = k.split(':').map(Number)
      out.push(placeOf(s, r, c))
    }
  return keysOf(out)
}
