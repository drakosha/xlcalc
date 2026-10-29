/**
 * Excel's Go To Special (`Session.getSpecialCells`, the set `bench/coverage/select/special`): the cells of a kind in a
 * range, a range of one cell standing for the used range (`usedArea`) and a larger one taken within it; the visible
 * cells of the range itself, or of the whole sheet for one cell. A blank cell
 * holds neither a value nor a formula; a constant is a value typed in, a formula is a cell with one, each by the kind
 * of its value (a formula answering empty text is text); a visible cell lies in no hidden row or column. The cells of
 * a spill past its formula are neither constants nor formulas, nor blank. The cells found are given as rectangles: the
 * runs of each row, a run joined to the one above it over the same columns.
 */
import type { CellValue, SpecialCells, SpecialValue } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { array, checkOneOf, object, onlyKeys } from '../input/basic.js'
import { lines, type Line } from '../layout/spans.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { cellValue } from '../result.js'
import { intersect, rectText, type Rect } from '../structure/area.js'
import { usedArea } from './extent.js'

const TYPES = ['blanks', 'constants', 'formulas', 'visible'] as const
const VALUES: readonly SpecialValue[] = ['numbers', 'text', 'logicals', 'errors']

/** The kind of cells, checked. */
export function checkSpecial(v: unknown): SpecialCells {
  const o = object(v, 'kind', "a kind of cells such as { type: 'constants', values: ['numbers'] }")
  checkOneOf(o.type, 'kind.type', TYPES)
  onlyKeys(o, 'kind', o.type === 'constants' || o.type === 'formulas' ? ['type', 'values'] : ['type'])
  if (o.values !== undefined) {
    const values = array(o.values, 'kind.values', 'kinds of values such as ["numbers", "text"]')
    values.forEach((x, i) => checkOneOf(x, `kind.values[${i}]`, VALUES))
    if (!values.length) throw new InputError('kind.values', 'expected at least one kind of values: numbers, text, logicals or errors')
  }
  return o as SpecialCells
}

/** The kind of a value. */
function valueKind(v: CellValue): SpecialValue | undefined {
  if (typeof v === 'number') return 'numbers'
  if (typeof v === 'string') return 'text'
  if (typeof v === 'boolean') return 'logicals'
  return v === null ? undefined : 'errors'
}

/** Cells as rectangles: each row's runs, a run joined to the rectangle above it over the same columns. */
function rectangles(rows: Map<number, number[]>): Rect[] {
  const done: Rect[] = []
  let open = new Map<string, Rect>()
  for (const r of [...rows.keys()].sort((a, b) => a - b)) {
    const next = new Map<string, Rect>()
    const cols = rows.get(r)!
    for (let i = 0; i < cols.length; ) {
      let j = i
      while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++
      const key = `${cols[i]}:${cols[j]}`
      const above = open.get(key)
      if (above && above.r2 === r - 1) {
        above.r2 = r
        open.delete(key)
        next.set(key, above)
      } else next.set(key, { r1: r, r2: r, c1: cols[i], c2: cols[j] })
      i = j + 1
    }
    done.push(...open.values())
    open = next
  }
  return [...done, ...open.values()].sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1)
}

/** The parts of the span `lo`..`hi` of an axis outside its hidden lines. */
function shownSpans(list: Line<{ hidden?: boolean }>[], lo: number, hi: number): [number, number][] {
  const out: [number, number][] = []
  let at = lo
  for (const [a, b, e] of list) {
    if (!e.hidden || b < at || a > hi) continue
    if (a > at) out.push([at, a - 1])
    at = b + 1
  }
  if (at <= hi) out.push([at, hi])
  return out
}

export function specialCells(state: EngineState, s: number, x: Rect, kind: SpecialCells): string[] {
  const sheet = state.input.sheets[s]
  const one = x.r1 === x.r2 && x.c1 === x.c2
  // the visible cells of the range, or of the sheet for one cell, by the spans of rows and columns not hidden
  if (kind.type === 'visible') {
    const area = one ? { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 } : x
    const cols = shownSpans(lines(sheet.columns, 'c'), area.c1, area.c2)
    return shownSpans(lines(sheet.rows, 'r'), area.r1, area.r2).flatMap(([r1, r2]) => cols.map(([c1, c2]) => rectText({ r1, r2, c1, c2 })))
  }
  const used = usedArea(state, s)
  const area = used && (one ? used : intersect(x, used))
  if (!area) return []
  const values = kind.type === 'constants' || kind.type === 'formulas' ? new Set(kind.values ?? VALUES) : undefined
  const rows = new Map<number, number[]>()
  for (let r = area.r1; r <= area.r2; r++)
    for (let c = area.c1; c <= area.c2; c++) {
      const own = sheet.cells[formatAddress(r, c)]
      const v = cellValue(state, s, r, c)
      const formula = own?.formula !== undefined
      let found: boolean
      if (kind.type === 'blanks') found = !formula && v === null
      else {
        const k = valueKind(v)
        found = (kind.type === 'formulas' ? formula : !formula && own?.value !== undefined && own.value !== null) && k !== undefined && values!.has(k)
      }
      if (found) (rows.get(r) ?? rows.set(r, []).get(r)!).push(c)
    }
  return rectangles(rows).map(rectText)
}
