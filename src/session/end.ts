/**
 * Where Control and an arrow key go from a cell (`Session.getEnd`), as Excel's `Range.End`: along the cell's column or
 * row, by the values the engine holds (a formula holds whatever it shows, a spill its stored and filled cells), hidden
 * lines passed over; it reads the line's cells up to where the sheet's values end, never the whole sheet.
 */
import type { EndDirection } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError, what } from '../input-error.js'
import { hiddenSpansOf } from '../layout/hidden.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { lowerBound } from '../sorted.js'

const DIRECTIONS: readonly string[] = ['up', 'down', 'left', 'right']

export function checkDirection(d: unknown): EndDirection {
  if (!DIRECTIONS.includes(d as string)) throw new InputError('direction', `expected up, down, left or right, got ${what(d)}`)
  return d as EndDirection
}

/** The cell `Range.End` comes to from row `r` and column `c` of sheet `s`, as [row, column]. */
export function endOf(state: EngineState, s: number, r: number, c: number, d: EndDirection): [number, number] {
  const vertical = d === 'up' || d === 'down'
  const step = d === 'down' || d === 'right' ? 1 : -1
  const sheet = state.input.sheets[s]
  const data = state.ctx.sheets[s]
  const max = (vertical ? MAX_ROWS : MAX_COLS) - 1
  const spans = vertical ? hiddenSpansOf(sheet.rows, 'r') : hiddenSpansOf(sheet.columns, 'c')
  // the cells of a large spill past its stored ones, as spans of the line
  const fills = data.fills.flatMap((f) => (vertical ? (f.c1 <= c && c <= f.c2 ? [[f.r1, f.r2]] : []) : f.r1 <= r && r <= f.r2 ? [[f.c1, f.c2]] : []))
  const col = data.cols[c]
  const held = (i: number) => (vertical ? col?.get(i) : data.cols[i]?.get(r)) != null || fills.some(([lo, hi]) => lo <= i && i <= hi)
  // past `last` the line holds nothing
  const last = fills.reduce((m, [, hi]) => Math.max(m, hi), (vertical ? (col?.length ?? 0) : data.cols.length) - 1)
  // line `i` if it is shown, else the first shown one beyond it going `by`; -1 past the sheet's edge
  const shown = (i: number, by: number) => {
    const k = lowerBound(spans, i, (x) => x[1])
    if (k < spans.length && spans[k][0] <= i) i = by > 0 ? spans[k][1] + 1 : spans[k][0] - 1
    return i < 0 || i > max ? -1 : i
  }
  const at = vertical ? r : c
  const next = shown(at + step, step)
  if (next < 0) return [r, c]
  const to = (i: number): [number, number] => (vertical ? [i, c] : [r, i])
  // from a cell with a value next to another: the last of the block
  if (held(at) && held(next)) {
    let i = next
    for (let j = shown(i + step, step); j >= 0 && held(j); j = shown(j + step, step)) i = j
    return to(i)
  }
  // else the next cell with a value, or the last shown line before the edge
  for (let j = step > 0 ? next : shown(Math.min(next, last), -1); j >= 0 && (step < 0 || j <= last); j = shown(j + step, step)) if (held(j)) return to(j)
  return to(step > 0 ? shown(max, -1) : shown(0, 1))
}
