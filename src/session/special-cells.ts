/**
 * Excel's Go To Special (`Session.getSpecialCells`, the set `bench/coverage/select/special`): the cells of a kind in a
 * range, a range of one cell standing for the used range (`usedArea`) and a larger one taken within it. A blank cell
 * holds neither a value nor a formula; a constant is a value typed in, a formula is a cell with one, each by the kind
 * of its value (a formula answering empty text is text); a visible cell lies in no hidden row or column. The cells of
 * a spill past its formula are neither constants nor formulas, nor blank. The cells found are given as rectangles: the
 * runs of each row, a run joined to the one above it over the same columns.
 */
import type { CellValue, SpecialCells, SpecialValue } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { array, checkOneOf, object, onlyKeys } from '../input/basic.js'
import { sheetLines } from '../layout/lines.js'
import { formatAddress } from '../model.js'
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

export function specialCells(state: EngineState, s: number, x: Rect, kind: SpecialCells): string[] {
  const used = usedArea(state, s)
  const area = used && (x.r1 === x.r2 && x.c1 === x.c2 ? used : intersect(x, used))
  if (!area) return []
  const sheet = state.input.sheets[s]
  const lines = sheetLines(sheet, state)
  const values = kind.type === 'constants' || kind.type === 'formulas' ? new Set(kind.values ?? VALUES) : undefined
  const rows = new Map<number, number[]>()
  for (let r = area.r1; r <= area.r2; r++) {
    const hiddenRow = lines.row(r)?.hidden
    for (let c = area.c1; c <= area.c2; c++) {
      let found: boolean
      if (kind.type === 'visible') found = !hiddenRow && !lines.column(c)?.hidden
      else {
        const own = sheet.cells[formatAddress(r, c)]
        const v = cellValue(state, s, r, c)
        const formula = own?.formula !== undefined
        if (kind.type === 'blanks') found = !formula && v === null
        else {
          const k = valueKind(v)
          found = (kind.type === 'formulas' ? formula : !formula && own?.value !== undefined && own.value !== null) && k !== undefined && values!.has(k)
        }
      }
      if (found) (rows.get(r) ?? rows.set(r, []).get(r)!).push(c)
    }
  }
  return rectangles(rows).map(rectText)
}
