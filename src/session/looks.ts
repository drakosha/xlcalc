/**
 * What a session's cells show, compared before and after a change that recomputes nothing (a cell's format, a
 * column's width, a row's or a column's format, merged areas): the cells whose shown text or colour changed.
 */
import type { CellValue, LayoutEdit } from '../contract.js'
import { toCellValue, valueAt, type EngineState } from '../engine.js'
import { parseMerge } from '../layout/merges.js'
import { parseSpan } from '../layout/spans.js'
import { parseAddress } from '../model.js'
import { outValue } from '../result.js'
import { cellLooks, type Looks } from '../shown.js'
import { sheetKey } from '../syntax.js'

/** A cell's value as a result shows it. */
export const cellValue = (s: EngineState, i: number, r: number, c: number): CellValue => outValue(toCellValue(valueAt(s, i, r, c)))

/** What a cell shows, as one comparable string. */
export const shownKey = (look: Looks, s: EngineState, i: number, r: number, c: number): string => {
  const d = look(i, r, c, cellValue(s, i, r, c))
  return d.color === undefined ? d.text : `${d.text}\u0000${d.color}`
}

/** What the cells `keys` (`s:r:c`) show now. */
export function snapshot(s: EngineState, keys: Iterable<string>): Map<string, string> {
  const look = cellLooks(s)
  const out = new Map<string, string>()
  for (const k of keys) {
    const [i, r, c] = k.split(':').map(Number)
    out.set(k, shownKey(look, s, i, r, c))
  }
  return out
}

/** The cells of `before` that show something else now. */
export function shownChanges(s: EngineState, before: Map<string, string>): string[] {
  const now = snapshot(s, before.keys())
  return [...before].filter(([k, was]) => now.get(k) !== was).map(([k]) => k)
}

/** The cells a layout edit can change the look of: the cells of its columns or rows, the top-left cells of merged areas old and new. */
export function layoutCells(s: EngineState, edits: LayoutEdit[]): Set<string> {
  const out = new Set<string>()
  for (const e of edits) {
    const i = s.sheetIdx.get(sheetKey(e.sheet))!
    const cols = s.ctx.sheets[i].cols
    if ('merges' in e) {
      for (const m of [...(s.input.sheets[i].merges ?? []), ...e.merges]) {
        const x = parseMerge(m)!
        out.add(`${i}:${x.r1}:${x.c1}`)
      }
    } else if ('columns' in e) {
      const [lo, hi] = parseSpan(e.columns, 'c')!
      for (let c = lo; c <= Math.min(hi, cols.length - 1); c++) cols[c]?.forEach((_, r) => out.add(`${i}:${r}:${c}`))
    } else {
      const [lo, hi] = parseSpan(e.rows, 'r')!
      cols.forEach((col, c) => {
        for (let r = lo; r <= Math.min(hi, col.length - 1); r++) if (col[r] !== undefined) out.add(`${i}:${r}:${c}`)
      })
    }
  }
  return out
}

/** The keys of edited cells (`sheet`, `cell`). */
export const editedCells = (s: EngineState, edits: { sheet: string; cell: string }[]): string[] =>
  edits.map((e) => {
    const p = parseAddress(e.cell)
    return `${s.sheetIdx.get(sheetKey(e.sheet))}:${p.row}:${p.col}`
  })
