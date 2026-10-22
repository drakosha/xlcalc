/**
 * Array formulas of the old kind (`CellInput.array`, entered with Ctrl+Shift+Enter): the formula's answer fitted to
 * its area as Excel fits it, and the session's edits of part of an area refused, as Excel refuses them.
 */
import type { Arg } from '../args.js'
import { Arr, broadcastAt } from '../arrays/value.js'
import { arrayArea, crossing, holds } from '../array-areas.js'
import type { CellInput, CellReplace } from '../contract.js'
import type { EngineState } from '../engine.js'
import { allot } from '../exec/budget.js'
import { InputError } from '../input-error.js'
import { formatAddress } from '../model.js'
import { cellCount, meets, rectCells, type Rect } from '../structure/area.js'
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
  for (const { s, top, area: x } of arrayAreas(state)) {
    const touched: string[] = []
    for (const a of rectCells(x)) if (edited.has(`${s}:${a}`)) touched.push(a)
    const all = touched.length === cellCount(x)
    const area = state.input.sheets[s].cells[top].array!
    const again = touched.length === 1 && touched[0] === top && edited.get(`${s}:${top}`)?.array === area
    if (touched.length && !all && !again) return { s, top, cell: touched[0], area }
  }
  return null
}

/** The areas of the book's array formulas of the old kind, by sheet and top-left cell. */
function arrayAreas(state: EngineState): { s: number; top: string; area: Rect }[] {
  const out: { s: number; top: string; area: Rect }[] = []
  if (state.arrays) for (const f of state.formulas) if (f?.area) out.push({ s: f.s, top: formatAddress(f.r, f.c), area: f.area })
  return out
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
 * area must hold nothing else after the edits, cross no other array formula's area and, of several cells, no table. A book without such formulas, and edits that make none, are not looked at.
 */
export function checkArrayEdits(state: EngineState, edits: CellReplace[]) {
  const made = edits.filter((e) => e.input?.array !== undefined)
  if (!made.length && !state.arrays) return
  const at = (s: number, addr: string) => `${s}:${addr}`
  const edited = editedOf(state, edits)
  const p = partOf(state, edited)
  if (p) throw new InputError('edits', `${p.cell} is part of the array formula at ${state.sheetNames[p.s]}!${p.top} (${p.area}): Excel changes all of an array or none of it`)
  // a new area holds nothing else once the edits are applied, and crosses no other array's area
  const kept = made.length ? arrayAreas(state).filter((a) => !edited.has(at(a.s, a.top))) : []
  for (const e of made) {
    const s = state.sheetIdx.get(sheetKey(e.sheet))!
    const x = arrayArea(e.cell, e.input!.array!)
    if (!x || e.input!.formula === undefined) throw new InputError('edits', `${e.sheet}!${e.cell}: expected the area of the cell's array formula with the cell at its top-left, got ${JSON.stringify(e.input!.array)}`)
    for (let r = x.r1; r <= x.r2; r++)
      for (let c = x.c1; c <= x.c2; c++) {
        const addr = formatAddress(r, c)
        const k = at(s, addr)
        if (addr !== e.cell && holds(edited.has(k) ? (edited.get(k) ?? undefined) : state.input.sheets[s].cells[addr])) throw new InputError('edits', `${e.sheet}!${addr} is inside the area of the array formula at ${e.cell} (${e.input!.array}), so it cannot hold anything else`)
      }
    const table = (x.r2 > x.r1 || x.c2 > x.c1) && state.tables.list.find((t) => t.s === s && meets(t, x))
    if (table) throw new InputError('edits', `${e.sheet}!${e.cell}: the array formula over ${e.input!.array} overlaps the table ${table.name} (Excel does not allow array formulas of several cells in tables)`)
    for (const a of kept) {
      const why = a.s === s && crossing(x, `${e.sheet}!${e.cell}`, a.area, `${state.sheetNames[s]}!${a.top}`)
      if (why) throw new InputError('edits', why)
    }
    kept.push({ s, top: e.cell, area: x })
  }
}
