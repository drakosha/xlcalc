/**
 * Array formulas of the old kind (`CellInput.array`, entered with Ctrl+Shift+Enter): the formula's answer fitted to
 * its area as Excel fits it, and the session's edits of part of an area refused, as Excel refuses them.
 */
import type { Arg } from '../args.js'
import { Arr, broadcastAt } from '../arrays/value.js'
import { arrayArea, holds } from '../array-areas.js'
import type { CellInput, CellReplace } from '../contract.js'
import type { EngineState } from '../engine.js'
import { allot } from '../exec/budget.js'
import { InputError } from '../input-error.js'
import { formatAddress } from '../model.js'
import { cellCount, rectCells } from '../structure/area.js'
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
 * The first array formula of the old kind whose area edits of a session change in part, as Excel refuses it: every cell
 * of the area must be edited, unless only its top-left cell is, keeping an array formula over the same area; null when
 * none is. `edited`: the edits' inputs by `s:addr`.
 */
function partOf(state: EngineState, edited: Map<string, CellInput | null>): { s: number; top: string; cell: string; area: string } | null {
  for (const f of state.formulas) {
    if (!f?.area) continue
    const x = f.area
    const top = formatAddress(f.r, f.c)
    const touched: string[] = []
    for (const a of rectCells(x)) if (edited.has(`${f.s}:${a}`)) touched.push(a)
    const all = touched.length === cellCount(x)
    const area = state.input.sheets[f.s].cells[top].array!
    const again = touched.length === 1 && touched[0] === top && edited.get(`${f.s}:${top}`)?.array === area
    if (touched.length && !all && !again) return { s: f.s, top, cell: touched[0], area }
  }
  return null
}

const editedOf = (state: EngineState, edits: CellReplace[]) => new Map(edits.map((e) => [`${state.sheetIdx.get(sheetKey(e.sheet))!}:${e.cell}`, e.input]))

/** The array formula of the old kind whose area the edits change in part (`partOf`): its sheet, its cell and its area; null when none. */
export function arrayPart(state: EngineState, edits: CellReplace[]): { sheet: string; cell: string; area: string } | null {
  if (!state.arrays) return null
  const p = partOf(state, editedOf(state, edits))
  return p && { sheet: state.sheetNames[p.s], cell: p.top, area: p.area }
}

/**
 * Edits of a session that change part of an array formula's area throw `InputError` (`partOf`). A new array formula's
 * area must hold nothing else after the edits. A book without such formulas, and edits that make none, are not looked at.
 */
export function checkArrayEdits(state: EngineState, edits: CellReplace[]) {
  const made = edits.filter((e) => e.input?.array !== undefined)
  if (!made.length && !state.arrays) return
  const at = (s: number, addr: string) => `${s}:${addr}`
  const edited = editedOf(state, edits)
  const p = partOf(state, edited)
  if (p) throw new InputError('edits', `${p.cell} is part of the array formula at ${state.sheetNames[p.s]}!${p.top} (${p.area}): Excel changes all of an array or none of it`)
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
