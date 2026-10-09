/**
 * Array formulas of the old kind (`CellInput.array`, entered with Ctrl+Shift+Enter): the formula's answer fitted to
 * its area as Excel fits it, and the session's edits of part of an area refused, as Excel refuses them.
 */
import type { Arg } from '../args.js'
import { Arr, broadcastAt } from '../arrays/value.js'
import { arrayArea, holds } from '../array-areas.js'
import type { CellEdit } from '../contract.js'
import type { EngineState } from '../engine.js'
import { allot } from '../exec/budget.js'
import { InputError } from '../input.js'
import { formatAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { Areas, E, Rng, type V } from '../values.js'

/**
 * The answer as the area's array: a value in every cell, an array or a reference repeated along a dimension of one,
 * else #N/A past its end, cut past the area; an empty element is 0.
 */
export function fit(x: Arg, rows: number, cols: number, view: (x: Arg) => V | Arr): Arr {
  const a = x instanceof Areas ? E.VALUE : x instanceof Rng || x instanceof Arr ? view(x) : x
  allot(rows * cols)
  const v: V[] = new Array(rows * cols)
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const value = broadcastAt(a, r, c)
      v[r * cols + c] = value === undefined ? E.NA : value
    }
  return new Arr(rows, cols, rows, cols, v)
}

/**
 * Edits of a session that change part of an array formula's area throw `InputError`: every cell of the area must be
 * edited, unless only its top-left cell is, keeping an array formula over the same area. A new array formula's area
 * must hold nothing else after the edits. A book without such formulas, and edits that make none, are not looked at.
 */
export function checkArrayEdits(state: EngineState, edits: CellEdit[]) {
  const made = edits.filter((e) => e.input?.array !== undefined)
  const areas = state.formulas.filter((f) => f?.area)
  if (!areas.length && !made.length) return
  const at = (s: number, addr: string) => `${s}:${addr}`
  const edited = new Map(edits.map((e) => [at(state.sheetIdx.get(sheetKey(e.sheet))!, e.cell), e.input]))
  for (const f of areas) {
    const x = f!.area!
    const top = formatAddress(f!.r, f!.c)
    const touched: string[] = []
    for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) if (edited.has(at(f!.s, formatAddress(r, c)))) touched.push(formatAddress(r, c))
    const all = touched.length === (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)
    const again = touched.length === 1 && touched[0] === top && edited.get(at(f!.s, top))?.array === state.input.sheets[f!.s].cells[top].array
    if (touched.length && !all && !again) throw new InputError('edits', `${touched[0]} is part of the array formula at ${state.sheetNames[f!.s]}!${top} (${state.input.sheets[f!.s].cells[top].array}): Excel changes all of an array or none of it`)
  }
  // a new area holds nothing else once the edits are applied
  for (const e of made) {
    const s = state.sheetIdx.get(sheetKey(e.sheet))!
    const x = arrayArea(e.cell, e.input!.array!)
    if (!x || e.input!.formula === undefined) throw new InputError('edits', `${e.sheet}!${e.cell}: expected the area of the cell's array formula with the cell at its top-left, got ${JSON.stringify(e.input!.array)}`)
    for (let r = x.r1; r <= x.r2; r++)
      for (let c = x.c1; c <= x.c2; c++) {
        const addr = formatAddress(r, c)
        const k = at(s, addr)
        if (addr !== e.cell && holds(edited.has(k) ? (edited.get(k) ?? undefined) : state.input.sheets[s].cells[addr])) throw new InputError('edits', `${e.sheet}!${addr} is in the area of the array formula at ${e.cell} (${e.input!.array}), which holds nothing else`)
      }
  }
}
